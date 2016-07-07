using System;
using System.Collections;
using System.IO;
using System.Security.AccessControl;


namespace HyperVSamples
{
    public class WinHelper
    {
        public static Object _lockw = new Object();

        public static char getFreeDisks()
        {
            //List<String> unusedDrives = new List<String>();
            ArrayList driveLetters = new ArrayList(26); // Allocate space for alphabet
            for (int i = 65; i < 91; i++) // increment from ASCII values for A-Z
            {
                driveLetters.Add(Convert.ToChar(i)); // Add uppercase letters to possible drive letters
            }

            foreach (string drive in Directory.GetLogicalDrives())
            {
                driveLetters.Remove(drive[0]); // removed used drive letters from possible drive letters
            }
            foreach (char element in driveLetters)
            {
                char drive = element;
                if (!Program.vdipathToDrive.ContainsValue(drive) )
                    //return drive;
            }
            //return 'Z';
        }

        public static bool copyFolderToVhd(string srcPath, string drive)
        {
            if (!Directory.Exists(srcPath))
            {
                Console.WriteLine("Source path '" + srcPath + "' doesnt exist, cant copy to the drive");
                return false;
            }

            string destPath = drive + ":\\";
            CopyFolderContents(srcPath, destPath);

            return true;
        }

        public static bool CopyFolderContents(string SourcePath, string DestinationPath)
        {
            SourcePath = SourcePath.EndsWith(@"\") ? SourcePath : SourcePath + @"\";
            DestinationPath = DestinationPath.EndsWith(@"\") ? DestinationPath : DestinationPath + @"\";

            try
            {
                if (Directory.Exists(SourcePath))
                {
                    if (Directory.Exists(DestinationPath) == false)
                    {
                        Directory.CreateDirectory(DestinationPath);
                    }

                    foreach (string files in Directory.GetFiles(SourcePath))
                    {
                        FileInfo fileInfo = new FileInfo(files);
                        // Console.WriteLine("Copying file '" + DestinationPath + "'");
                        // fileInfo.MoveTo(string.Format(@"{0}\{1}", DestinationPath, fileInfo.Name));
                        fileInfo.CopyTo(string.Format(@"{0}\{1}", DestinationPath, fileInfo.Name));
                    }

                    foreach (string drs in Directory.GetDirectories(SourcePath))
                    {
                        DirectoryInfo directoryInfo = new DirectoryInfo(drs);
                        if (CopyFolderContents(drs, DestinationPath + directoryInfo.Name) == false)
                        {
                            return false;
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public static bool fixPermission(string location)
        {
            try
            {
                const string filePermissionsGroup = "Users";
                if (SetFilePermissions(filePermissionsGroup, FileSystemRights.FullControl, Path.Combine(location)))
                {
                    return true;
                }
                else
                {
                    Console.WriteLine("Failed to set file permission for '" + location + "'");
                    return false;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Exception trying to set file permission for '" + location + "'");
                return false;
            }
        }

        public static bool SetFilePermissions(string userOrGroupName, FileSystemRights rights, string directoryPath)
        {
            var accessRule = new FileSystemAccessRule(userOrGroupName, rights, InheritanceFlags.None, PropagationFlags.NoPropagateInherit, AccessControlType.Allow);

            var directoryInfo = new DirectoryInfo(directoryPath);
            var directorySecurity = directoryInfo.GetAccessControl(AccessControlSections.Access);

            var isModifyAccessRuleWorked = false;
            directorySecurity.ModifyAccessRule(AccessControlModification.Set, accessRule, out isModifyAccessRuleWorked);

            if (!isModifyAccessRuleWorked)
            {
                return false;
            }

            var iFlags = InheritanceFlags.ObjectInherit;
            iFlags = InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit;

            accessRule = new FileSystemAccessRule(userOrGroupName, rights, iFlags, PropagationFlags.InheritOnly, AccessControlType.Allow);
            isModifyAccessRuleWorked = false;
            directorySecurity.ModifyAccessRule(AccessControlModification.Add, accessRule, out isModifyAccessRuleWorked);

            if (!isModifyAccessRuleWorked)
            {
                return false;
            }

            directoryInfo.SetAccessControl(directorySecurity);
            return true;
        }

        public static bool verifyVHD(string vhdPath, string drive)
        {
            try
            {
                vhdHelper.AttachVirtualHardDisk(vhdPath);
                // drive = vhdHelper.assignDriveLetter(uniqueId);

                Console.WriteLine("Successfully attached the final vhd at '" + vhdPath + "' to drive '" + drive + "'");
                string[] files = Directory.GetFiles(drive + ":\\", "*.*", SearchOption.TopDirectoryOnly);

                // Display all the files.
                foreach (string file in files)
                {
                    continue;
                }
                if (Directory.Exists(drive + ":\\asgard"))
                {
                    Console.WriteLine("Successfully verfied the contents of the final vhd at '" + vhdPath + "' to drive '" + drive + "'");
                    vhdHelper.DetachVirtualHardDisk(vhdPath);
                    return true;
                }
                else
                {
                    Console.WriteLine("Failed to verify contents of the final vhd at '" + vhdPath + "' to drive '" + drive + "'");
                    vhdHelper.DetachVirtualHardDisk(vhdPath);
                    return false;
                }
            }
            catch(Exception e)
            {
                Console.WriteLine("Failed to attach the final vhd at '" + vhdPath + "' to drive '" + drive + "'");
                vhdHelper.DetachVirtualHardDisk(vhdPath);
                return false;
            }
        }
    }
}
