using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using dnlib.DotNet;
using dnlib.DotNet.Emit;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using System.Net;
using System.Diagnostics;

namespace UnityModManagerNet.Installer
{
    public partial class UnityModManagerForm : Form
    {
        public UnityModManagerForm()
        {
            InitializeComponent();
            Init();
            InitPageMods();
        }

        public static UnityModManagerForm instance = null;

        static Config config = null;
        static Param param = null;
        static Version version = null;
        static string currentGamePath = null;
        static string currentManagedPath = null;

        GameInfo selectedGame => (GameInfo)gameList.SelectedItem;
        ModInfo selectedMod => listMods.SelectedItems.Count > 0 ? mods.Find(x => x.DisplayName == listMods.SelectedItems[0].Text) : null;

        private void Init()
        {
            FormBorderStyle = FormBorderStyle.FixedDialog;
            instance = this;

            Log.Init();

            var modManagerType = typeof(UnityModManager);
            var modManagerDef = ModuleDefMD.Load(modManagerType.Module);
            var modManager = modManagerDef.Types.First(x => x.Name == modManagerType.Name);
            var versionString = modManager.Fields.First(x => x.Name == nameof(UnityModManager.version)).Constant.Value.ToString();
            currentVersion.Text = versionString;
            version = Utils.ParseVersion(versionString);

            config = Config.Load();
            param = Param.Load();

            if (config != null && config.GameInfo != null && config.GameInfo.Length > 0)
            {
                gameList.Items.AddRange(config.GameInfo);

                GameInfo selected = null;
                if (!string.IsNullOrEmpty(param.LastSelectedGame))
                {
                    selected = config.GameInfo.FirstOrDefault(x => x.Name == param.LastSelectedGame);
                }
                gameList.SelectedItem = selected ?? config.GameInfo.First();
            }
            else
            {
                InactiveForm();
                Log.Print($"解析文件错误 ： '{Config.filename}'.");
                return;
            }

            CheckLastVersion();
        }

        private void UnityModLoaderForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Properties.Settings.Default.Save();
            param.Save();
        }

        private void InactiveForm()
        {
            btnInstall.Enabled = false;
            btnRemove.Enabled = false;
            tabControl.TabPages[1].Enabled = false;
        }

        private bool IsValid(GameInfo gameIfno)
        {
            if (selectedGame == null)
            {
                Log.Print("选择一个游戏.");
                return false;
            }

            string output = "";
            foreach (var field in typeof(GameInfo).GetFields())
            {
                if (!field.IsStatic && field.IsPublic)
                {
                    var value = field.GetValue(gameIfno);
                    if (value == null || value.ToString() == "")
                    {
                        output += $"属性值 '{field.Name}' 为空. ";
                    }
                }
            }
            if (!string.IsNullOrEmpty(output))
            {
                Log.Print(output);
                return false;
            }

            return true;
        }

        private void CheckState()
        {
            if (!IsValid(selectedGame))
            {
                InactiveForm();
                return;
            }

            btnInstall.Text = "安装";

            currentGamePath = "";
            if (!param.ExtractGamePath(selectedGame.Name, out var result) || !Directory.Exists(result))
            {
                result = FindGameFolder(selectedGame.Folder);
                if (string.IsNullOrEmpty(result))
                {
                    InactiveForm();
                    btnOpenFolder.ForeColor = System.Drawing.Color.FromArgb(192, 0, 0);
                    btnOpenFolder.Text = "选择游戏安装目录";
                    folderBrowserDialog.SelectedPath = null;
                    Log.Print($"游戏安装目录 '{selectedGame.Folder}' 无法找到.");
                    return;
                }
                Log.Print($"游戏安装目录 被检测为 '{result}'.");
                param.SaveGamePath(selectedGame.Name, result);
            }
            else
            {
                Log.Print($"游戏安装目录 被设置为 '{result}'.");
            }
            currentGamePath = result;
            btnOpenFolder.ForeColor = System.Drawing.Color.Black;
            btnOpenFolder.Text = new DirectoryInfo(result).Name;
            folderBrowserDialog.SelectedPath = currentGamePath;
            currentManagedPath = FindManagedFolder(currentGamePath);

            var assemblyPath = Path.Combine(currentManagedPath, selectedGame.AssemblyName);

            ModuleDefMD assembly = null;

            if (File.Exists(assemblyPath))
            {
                try
                {
                    if (this.selectedGame.Name == "The Scroll Of Taiwu")
                    {
                        string de_ass = assemblyPath + ".de.dll";
                        decodeass(assemblyPath, de_ass);
                        assembly = ModuleDefMD.Load(File.ReadAllBytes(de_ass));
                    }
                    else
                    {
                        assembly = ModuleDefMD.Load(File.ReadAllBytes(assemblyPath));
                    }
                        
                    
                }
                catch (Exception e)
                {
                    InactiveForm();
                    Log.Print(e.Message);
                    return;
                }
            }
            else
            {
                InactiveForm();
                Log.Print($"'{selectedGame.AssemblyName}' not found.");
                return;
            }

            tabControl.TabPages[1].Enabled = true;

            var modManagerType = typeof(UnityModManager);
            var modManagerDefInjected = assembly.Types.FirstOrDefault(x => x.Name == modManagerType.Name);
            if (modManagerDefInjected != null)
            {
                btnInstall.Text = "更新";
                btnInstall.Enabled = false;
                btnRemove.Enabled = true;

                var versionString = modManagerDefInjected.Fields.First(x => x.Name == nameof(UnityModManager.version)).Constant.Value.ToString();
                var version2 = Utils.ParseVersion(versionString);
                installedVersion.Text = versionString;
                if (version != version2)
                { 
                    btnInstall.Enabled = true;
                }
            }
            else
            {
                installedVersion.Text = "-";
                btnInstall.Enabled = true;
                btnRemove.Enabled = false;
            }

            if (this.selectedGame.Name == "The Scroll Of Taiwu")
            {
                string de_ass = assemblyPath + ".de.dll";
                File.Delete(de_ass);
            }

        }


        private bool decodeass(string assfile, string de_assfile)
        {
            if (System.IO.File.Exists(assfile))
            {
                byte[] fd = System.IO.File.ReadAllBytes(assfile);

                    fd = Xxtea.XXTEA.Decrypt(fd, Xxtea.nhelper.dd("E93125554AC396980E3F3F57F70BE8CAADF3E9F7BCC7BC56DF8E124F06A20DEC850A631886F98A4E08CB976A52017EA33F1187563439BA19D7D77520EF4FD9C5"));
                    System.IO.File.WriteAllBytes(de_assfile, fd);
                    Log.Print($"'{selectedGame.AssemblyName}' unpacked.");
                return true;
            }
            return false;
        }

        private bool encodeass(string assfile, string en_assfile)
        {
            if (System.IO.File.Exists(assfile))
            {
                byte[] fd = System.IO.File.ReadAllBytes(assfile);

                fd = Xxtea.XXTEA.Encrypt(fd, System.Text.Encoding.ASCII.GetBytes("43062619851001801xshipzheng12a8294639u5238798jie1986929conch"));
                System.IO.File.WriteAllBytes(en_assfile, fd);
                Log.Print($"'{selectedGame.AssemblyName}' packed.");
                return true;


            }
            return false;
        }

        private string FindGameFolder(string str)
        {
            string[] disks = new string[] { @"C:\", @"D:\", @"E:\", @"F:\" };
            string[] roots = new string[] { "Games", "Program files", "Program files (x86)", "" };
            string[] folders = new string[] { @"Steam\SteamApps\common", "" };
            foreach (var disk in disks)
            {
                foreach (var root in roots)
                {
                    foreach (var folder in folders)
                    {
                        var path = Path.Combine(disk, root);
                        path = Path.Combine(path, folder);
                        path = Path.Combine(path, str);
                        if (Directory.Exists(path))
                        {
                            return path;
                        }
                    }
                }
            }
            return null;
        }

        private string FindManagedFolder(string str)
        {
            var regex = new Regex(".*_Data$");
            var dirictory = new DirectoryInfo(str);
            foreach (var dir in dirictory.GetDirectories())
            {
                var match = regex.Match(dir.Name);
                if (match.Success)
                {
                    var path = Path.Combine(str, $"{dir.Name}{Path.DirectorySeparatorChar}Managed");
                    if (Directory.Exists(path))
                    {
                        return path;
                    }
                }
            }
            return str;
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (Inject(Actions.Remove))
                btnInstall.Text = "安装";
        }

        private void btnInstall_Click(object sender, EventArgs e)
        {
            if (Inject(Actions.Install))
                btnInstall.Text = "更新";
        }

        private void btnDownloadUpdate_Click(object sender, EventArgs e)
        {
            var downloaderFile = "Downloader.exe";
            if (File.Exists(downloaderFile))
            {
                Process.Start(downloaderFile);
            }
        }

        private void btnOpenFolder_Click(object sender, EventArgs e)
        {
            var result = folderBrowserDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                //inputLog.Clear();
                param.SaveGamePath(selectedGame.Name, folderBrowserDialog.SelectedPath);
                CheckState();
            }
        }

        private void gameList_Changed(object sender, EventArgs e)
        {
            //inputLog.Clear();
            var selected = (GameInfo)((ComboBox)sender).SelectedItem;
            if (selected != null)
                Log.Print($"游戏变更为 '{selected.Name}'.");

            param.LastSelectedGame = selected.Name;

            CheckState();

        }

        enum Actions
        {
            Install,
            Remove
        };

        private bool Inject(Actions action, ModuleDefMD assembly = null, bool save = true)
        {
            var assemblyPath = Path.Combine(currentManagedPath, selectedGame.AssemblyName);
            var backupAssemblyPath = $"{assemblyPath}.backup";

            if (File.Exists(assemblyPath))
            {
                if (assembly == null)
                {
                    try
                    {
                        if (this.selectedGame.Name == "The Scroll Of Taiwu")
                        {
                            string de_ass = assemblyPath + ".de.dll";
                            decodeass(assemblyPath, de_ass);
                            assembly = ModuleDefMD.Load(File.ReadAllBytes(de_ass));
                        }
                        else
                        {
                            assembly = ModuleDefMD.Load(File.ReadAllBytes(assemblyPath));
                        }

                        
                    }
                    catch (Exception e)
                    {
                        Log.Print(e.Message);
                        return false;
                    }
                }

                string className = null;
                string methodName = null;
                string placeType = null;

                var pos = selectedGame.PatchTarget.LastIndexOf('.');
                if (pos != -1)
                {
                    className = selectedGame.PatchTarget.Substring(0, pos);

                    var pos2 = selectedGame.PatchTarget.LastIndexOf(':');
                    if (pos2 != -1)
                    {
                        methodName = selectedGame.PatchTarget.Substring(pos + 1, pos2 - pos - 1);
                        placeType = selectedGame.PatchTarget.Substring(pos2 + 1).ToLower();

                        if (placeType != "after" && placeType != "before")
                            Log.Print($"Parameter '{placeType}' in '{selectedGame.PatchTarget}' is unknown.");
                    }
                    else
                    {
                        methodName = selectedGame.PatchTarget.Substring(pos + 1);
                    }

                    if (methodName == "ctor")
                        methodName = ".ctor";
                }
                else
                {
                    Log.Print($"Function name error '{selectedGame.PatchTarget}'.");
                    return false;
                }

                var targetClass = assembly.Types.FirstOrDefault(x => x.FullName == className);
                if (targetClass == null)
                {
                    Log.Print($"Class '{className}' not found.");
                    return false;
                }

                var targetMethod = targetClass.Methods.FirstOrDefault(x => x.Name == methodName);
                if (targetMethod == null)
                {
                    Log.Print($"Method '{methodName}' not found.");
                    return false;
                }

                var modManagerType = typeof(UnityModManager);

                switch (action)
                {
                    case Actions.Install:
                        try
                        {
                            Log.Print($"Backup for '{selectedGame.AssemblyName}'.");
                            File.Copy(assemblyPath, backupAssemblyPath, true);

                            CopyLibraries();

                            var modsPath = Path.Combine(currentGamePath, selectedGame.ModsDirectory);
                            if (!Directory.Exists(modsPath))
                            {
                                Directory.CreateDirectory(modsPath);
                            }

                            var typeInjectorInstalled = assembly.Types.FirstOrDefault(x => x.Name == modManagerType.Name);
                            if (typeInjectorInstalled != null)
                            {
                                if (!Inject(Actions.Remove, assembly, false))
                                {
                                    Log.Print("Installation failed. Can't uninstall the previous version.");
                                    return false;
                                }
                            }

                            Log.Print("Applying patch...");
                            var modManagerDef = ModuleDefMD.Load(modManagerType.Module);
                            var modManager = modManagerDef.Types.First(x => x.Name == modManagerType.Name);
                            var modManagerModsDir = modManager.Fields.First(x => x.Name == nameof(UnityModManager.modsDirname));
                            modManagerModsDir.Constant.Value = selectedGame.ModsDirectory;
                            var modManagerModInfo = modManager.Fields.First(x => x.Name == nameof(UnityModManager.infoFilename));
                            modManagerModInfo.Constant.Value = selectedGame.ModInfo;
                            modManagerDef.Types.Remove(modManager);
                            assembly.Types.Add(modManager);

                            var instr = OpCodes.Call.ToInstruction(modManager.Methods.First(x => x.Name == nameof(UnityModManager.Start)));
                            if (string.IsNullOrEmpty(placeType) || placeType == "after")
                            {
                                targetMethod.Body.Instructions.Insert(targetMethod.Body.Instructions.Count - 1, instr);
                            }
                            else if (placeType == "before")
                            {
                                targetMethod.Body.Instructions.Insert(0, instr);
                            }

                            if (save)
                            {
                                if (this.selectedGame.Name == "The Scroll Of Taiwu")
                                {
                                    string de_ass = assemblyPath + ".de.dll";
                                    assembly.Write(de_ass);
                                    encodeass(de_ass, assemblyPath);
                                    File.Delete(de_ass);
                                }
                                else
                                {
                                    assembly.Write(assemblyPath);
                                }

                                
                                Log.Print("安装成功.");
                            }

                            installedVersion.Text = currentVersion.Text;
                            btnInstall.Enabled = false;
                            btnRemove.Enabled = true;

                            return true;
                        }
                        catch (Exception e)
                        {
                            Log.Print(e.Message);
                            if (!File.Exists(assemblyPath))
                                RestoreBackup();
                        }

                        break;

                    case Actions.Remove:
                        try
                        {
                            var modManagerInjected = assembly.Types.FirstOrDefault(x => x.Name == modManagerType.Name);
                            if (modManagerInjected != null)
                            {
                                Log.Print("移除ModManager...");
                                var instr = OpCodes.Call.ToInstruction(modManagerInjected.Methods.First(x => x.Name == nameof(UnityModManager.Start)));
                                for (int i = 0; i < targetMethod.Body.Instructions.Count; i++)
                                {
                                    if (targetMethod.Body.Instructions[i].OpCode == instr.OpCode
                                        && targetMethod.Body.Instructions[i].Operand == instr.Operand)
                                    {
                                        targetMethod.Body.Instructions.RemoveAt(i);
                                        break;
                                    }
                                }

                                assembly.Types.Remove(modManagerInjected);

                                if (save)
                                {
                                    if (this.selectedGame.Name == "The Scroll Of Taiwu")
                                    {
                                        string de_ass = assemblyPath + ".de.dll";
                                        assembly.Write(de_ass);
                                        encodeass(de_ass, assemblyPath);
                                        File.Delete(de_ass);
                                    }
                                    else
                                    {
                                        assembly.Write(assemblyPath);
                                    }
                                    Log.Print("移除成功.");
                                }

                                installedVersion.Text = "-";
                                btnInstall.Enabled = true;
                                btnRemove.Enabled = false;
                            }

                            return true;
                        }
                        catch (Exception e)
                        {
                            Log.Print(e.Message);
                            if (!File.Exists(assemblyPath))
                                RestoreBackup();
                        }

                        break;
                }
            }
            else
            {
                Log.Print($"'{assemblyPath}' 无法找到.");
                return false;
            }

            return false;
        }

        private static bool RestoreBackup()
        {
            var assemblyPath = Path.Combine(currentManagedPath, instance.selectedGame.AssemblyName);
            var backupAssemblyPath = $"{assemblyPath}.backup";

            try
            {
                if (File.Exists(backupAssemblyPath))
                {
                    File.Copy(backupAssemblyPath, assemblyPath, true);
                    Log.Print("备份已恢复.");
                    return true;
                }
            }
            catch (Exception e)
            {
                Log.Print(e.Message);
            }

            return false;
        }

        private static void CopyLibraries()
        {
            string[] files = new string[]
            {
                //"0Harmony.dll",
                "0Harmony12.dll"
            };

            foreach (var file in files)
            {
                var path = Path.Combine(currentManagedPath, file);
                if (File.Exists(path))
                {
                    var source = new FileInfo(file);
                    var dest = new FileInfo(path);
                    if (dest.Length == source.Length)
                        continue;

                    File.Copy(path, $"{path}.backup", true);
                }

                File.Copy(file, path, true);
                Log.Print($"'{file}' 已经复制到游戏目录.");
            }
        }

        private void folderBrowserDialog_HelpRequest(object sender, EventArgs e)
        {
        }

        private void tabs_Changed(object sender, EventArgs e)
        {
            switch (tabControl.SelectedIndex)
            {
                case 1: // Mods
                    ReloadMods();
                    RefreshModList();
                    if (!repositories.ContainsKey(selectedGame))
                        CheckModUpdates();
                    break;
            }
        }
    }
}
