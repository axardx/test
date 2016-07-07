using HyperVSamples;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Security.AccessControl;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;

namespace HyperVSamplesTests
{
    [TestClass()]
    public class WinHelperTests
    {
        [TestMethod()]
        public void WinHelper_getFreeDisks_CheckReturnDriveLetter_ReturnDriveLetterZ()
        {
            Program.vdipathToDrive = new Dictionary<string, char>();

            for (int i = 65; i < 91; i++) // increment from ASCII values for A-Z
            {
                Program.vdipathToDrive[i.ToString()] = Convert.ToChar(i); // Add uppercase letters to possible drive letters
            }

            char drive = WinHelper.getFreeDisks();
            Assert.AreEqual('Z', drive);

            Program.vdipathToDrive.Clear();
        }

        [TestMethod()]
        public void WinHelper_copyFolderToVhd_CheckFolderExistInVHD_IsExist()
        {
            string srcPath = "C:\\test";
            string vdiPath = "C:\\test_copy.vhdx";
            UInt64 vdiSize = 2;
            char drive = 'Q';

            // create disk
            vhdHelper.CreateFixedVirtualHardDisk(vdiPath, vdiSize);
            // attach disk 
            vhdHelper.AttachVirtualHardDisk(vdiPath);
            // initialize and partition disk
            UInt32 diskNumber = Convert.ToUInt32(vhdHelper.partitionWMI(vdiPath, drive));
            // format disk
            vhdHelper.formatWMI(drive.ToString());

            WinHelper.copyFolderToVhd(srcPath, drive.ToString());

            Assert.AreEqual("True", Directory.Exists(drive.ToString() + ":\\asgard").ToString());
            Assert.AreEqual("True", File.Exists(drive.ToString() + ":\\New Text Document.txt").ToString());
            Assert.AreEqual("True", File.Exists(drive.ToString() + ":\\Thread Group.jmx").ToString());

            // detach and delete the file
            vhdHelper.DetachVirtualHardDisk(vdiPath);
            if (File.Exists(vdiPath))
                File.Delete(vdiPath);
        }

        [TestMethod()]
        public void WinHelper_copyFolderToVhd_CheckFolderExistInVHD_NotExist()
        {
            string srcPath = "C:\\test_12345";
            string driveletter = null;

            Assert.AreEqual("False", WinHelper.copyFolderToVhd(srcPath, driveletter).ToString());
        }

        [TestMethod()]
        public void WinHelper_CopyFolderContents_CheckFolderExistDes_IsExist()
        {
            string srcPath = "C:\\test";
            string desPath = "C:\\test_des";

            Directory.CreateDirectory(desPath);

            WinHelper.CopyFolderContents(srcPath, desPath);

            Assert.AreEqual("True", Directory.Exists(desPath + "\\asgard").ToString());
            Assert.AreEqual("True", File.Exists(desPath + "\\New Text Document.txt").ToString());
            Assert.AreEqual("True", File.Exists(desPath + "\\Thread Group.jmx").ToString());

            Directory.Delete(desPath, true);
        }

        [TestMethod()]
        public void WinHelper_fixPermission_CheckReadWritePermission_IsGranted()
        {
            string FileLocation = "C:\\test\\asgard";

            WinHelper.fixPermission(FileLocation);

            FileIOPermission writePermission = new FileIOPermission(FileIOPermissionAccess.Write, FileLocation);
            Assert.AreEqual("True", SecurityManager.IsGranted(writePermission).ToString());

            FileIOPermission readPermission = new FileIOPermission(FileIOPermissionAccess.Read, FileLocation);
            Assert.AreEqual("True", SecurityManager.IsGranted(readPermission).ToString());
        }

        [TestMethod()]
        public void WinHelper_SetFilePermissions_CheckReadWritePermission_IsGranted()
        {
            string FileLocation = "C:\\test\\asgard";
            const string filePermissionsGroup = "Users";
            FileSystemRights right = FileSystemRights.FullControl;

            WinHelper.SetFilePermissions(filePermissionsGroup, right, FileLocation);

            FileIOPermission writePermission = new FileIOPermission(FileIOPermissionAccess.Write, FileLocation);
            Assert.AreEqual("True", SecurityManager.IsGranted(writePermission).ToString());

            FileIOPermission readPermission = new FileIOPermission(FileIOPermissionAccess.Read, FileLocation);
            Assert.AreEqual("True", SecurityManager.IsGranted(readPermission).ToString());

        }

        [TestMethod()]
        public void WinHelper_verifyVHD_CheckVerificationSucceed_Succeed()
        {
            string srcPath = "C:\\test";
            string vdiPath = "C:\\test_verify.vhdx";
            UInt64 vdiSize = 2;
            char drive = 'Z';

            // create disk
            vhdHelper.CreateFixedVirtualHardDisk(vdiPath, vdiSize);
            // attach disk 
            vhdHelper.AttachVirtualHardDisk(vdiPath);
            // initialize and partition disk
            UInt32 diskNumber = Convert.ToUInt32(vhdHelper.partitionWMI(vdiPath, drive));
            // format disk
            vhdHelper.formatWMI(drive.ToString());
            // copy file from source path to vhdx
            WinHelper.copyFolderToVhd(srcPath, drive.ToString());
            // defrag disk
            vhdHelper.defragVHD(drive.ToString());
            // detach disk
            vhdHelper.DetachVirtualHardDisk(vdiPath);

            Assert.AreEqual("True", WinHelper.verifyVHD(vdiPath, drive.ToString()).ToString());

            if (File.Exists(vdiPath))
                File.Delete(vdiPath);

        }

        [TestMethod()]
        public void WinHelper_verifyVHD_CheckVerificationSucceed_Fail()
        {
            string srcPath = "C:\\test2";
            string vdiPath = "C:\\test_verify.vhdx";
            UInt64 vdiSize = 2;
            char drive = 'Z';

            // create disk
            vhdHelper.CreateFixedVirtualHardDisk(vdiPath, vdiSize);
            // attach disk 
            vhdHelper.AttachVirtualHardDisk(vdiPath);
            // initialize and partition disk
            UInt32 diskNumber = Convert.ToUInt32(vhdHelper.partitionWMI(vdiPath, drive));
            // format disk
            vhdHelper.formatWMI(drive.ToString());
            // copy file from source path to vhdx
            WinHelper.copyFolderToVhd(srcPath, drive.ToString());
            // defrag disk
            vhdHelper.defragVHD(drive.ToString());
            // detach disk
            vhdHelper.DetachVirtualHardDisk(vdiPath);

            Assert.AreEqual("False", WinHelper.verifyVHD(vdiPath, drive.ToString()).ToString());

            if (File.Exists(vdiPath))
                File.Delete(vdiPath);

        }

    }
}
