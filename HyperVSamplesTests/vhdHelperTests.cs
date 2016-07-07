using Microsoft.VisualStudio.TestTools.UnitTesting;
using HyperVSamples;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Management;
using System.Threading;

namespace HyperVSamples.Tests
{
    [TestClass()]
    public class vhdHelperTests
    {

        [TestMethod()]
        public void vhdHelper_CreateFixedVirtualHardDisk_CheckVHDXExist_IsExist()
        {
            string vdiPath = "C:\\test_create.vhdx";
            UInt64 vdiSize = 2;

            vhdHelper.CreateFixedVirtualHardDisk(vdiPath, vdiSize);

            Assert.AreEqual("True", File.Exists(vdiPath).ToString());

            if (File.Exists(vdiPath))
                File.Delete(vdiPath);
        }

        [TestMethod()]
        public void vhdHelper_AttachVirtualHardDisk_CheckDiskExist_IsExist()
        {
            string vdiPath = "C:\\test_attach.vhdx";
            UInt64 vdiSize = 2;

            // create file
            vhdHelper.CreateFixedVirtualHardDisk(vdiPath, vdiSize);

            // Test if there exist a disk with a disk number and 0 partition
            vhdHelper.AttachVirtualHardDisk(vdiPath);

            string sqe = string.Format("SELECT * FROM MSFT_Disk");
            //string sqe = string.Format("SELECT * FROM MSFT_Disk WHERE Number={0}", diskNumber);
            string diskNumber = null;
            ManagementObject disk = null;

            using (ManagementObjectSearcher diskDrives = new ManagementObjectSearcher(@"\\.\Root\Microsoft\Windows\Storage", sqe))
            {
                ManagementObjectCollection moc = diskDrives.Get();
                foreach (ManagementObject diskObj in moc)
                {
                    if (diskObj["NumberOfPartitions"].Equals((uint)0))
                    {
                        disk = diskObj;
                        diskNumber = diskObj["Number"].ToString();
                    }
                }
            }
            Assert.IsNotNull(disk);
            Assert.IsNotNull(diskNumber);


            // detach and delete file
            vhdHelper.DetachVirtualHardDisk(vdiPath);
            if (File.Exists(vdiPath))
                File.Delete(vdiPath);
        }

        [TestMethod()]
        public void vhdHelper_DetachVirtualHardDisk_CheckDiskNonExist_NonExist()
        {

            string vdiPath = "C:\\test_detach.vhdx";
            UInt64 vdiSize = 2;

            // create file
            vhdHelper.CreateFixedVirtualHardDisk(vdiPath, vdiSize);

            // attach file 
            vhdHelper.AttachVirtualHardDisk(vdiPath);

            // Test if there exist a disk with a disk number and 0 partition
            vhdHelper.DetachVirtualHardDisk(vdiPath);
            Thread.Sleep(1000);
            string sqe = string.Format("SELECT * FROM MSFT_Disk");
            //string sqe = string.Format("SELECT * FROM MSFT_Disk WHERE Number={0}", diskNumber);
            string diskNumber = null;
            ManagementObject disk = null;

            using (ManagementObjectSearcher diskDrives = new ManagementObjectSearcher(@"\\.\Root\Microsoft\Windows\Storage", sqe))
            {
                ManagementObjectCollection moc = diskDrives.Get();
                foreach (ManagementObject diskObj in moc)
                {
                    if (diskObj["NumberOfPartitions"].Equals((uint)0))
                    {
                        disk = diskObj;
                        diskNumber = diskObj["Number"].ToString();
                    }
                }
            }
            Assert.IsNull(disk);
            Assert.IsNull(diskNumber);

            // delete the file
            if (File.Exists(vdiPath))
                File.Delete(vdiPath);
        }

        [TestMethod()]
        public void vhdHelper_GetVirtualHardDiskSettingDataEmbeddedInstance_CheckStringReturned_NotNull()
        {
            string serverName = "\\";
            string result = null;
            string vdiPath = "C:\\test_getsettingdata.vhdx";
            UInt64 vdiSize = 2;

            ManagementScope scope = new ManagementScope(@"Root\Virtualization\V2", null);

            using (ManagementObject imageManagementService = vhdHelper.GetImageManagementService(scope))
            {
                result = vhdHelper.GetVirtualHardDiskSettingDataEmbeddedInstance(serverName, imageManagementService.Path.Path, vdiPath, vdiSize);
            }

            Assert.IsNotNull(result);
        }

        [TestMethod()]
        public void vhdHelper_initializeWMI_CheckPartitionStyle_IsMBR()
        {
            string vdiPath = "C:\\test_init.vhdx";
            UInt64 vdiSize = 2;

            // create file
            vhdHelper.CreateFixedVirtualHardDisk(vdiPath, vdiSize);

            // attach the file
            vhdHelper.AttachVirtualHardDisk(vdiPath);

            string sqe = string.Format("SELECT * FROM MSFT_Disk");
            string diskNumber = null;
            ManagementObject disk = null;

            using (ManagementObjectSearcher diskDrives = new ManagementObjectSearcher(@"\\.\Root\Microsoft\Windows\Storage", sqe))
            {
                ManagementObjectCollection moc = diskDrives.Get();
                foreach (ManagementObject diskObj in moc)
                {
                    if (diskObj["NumberOfPartitions"].Equals((uint)0))
                    {
                        disk = diskObj;
                        diskNumber = diskObj["Number"].ToString();
                    }
                }
            }

            // Test if partition style is MBR(1)
            vhdHelper.initializeWMI(disk, (uint)VirtualHardDiskPartitionStyle.MBR);

            string squ = string.Format("SELECT * FROM MSFT_Disk WHERE Number={0}", diskNumber);
            using (ManagementObjectSearcher diskDrives = new ManagementObjectSearcher(@"\\.\Root\Microsoft\Windows\Storage", sqe))
            {
                ManagementObjectCollection moc = diskDrives.Get();
                foreach (ManagementObject diskObj in moc)
                {
                    Assert.AreEqual("1", diskObj["PartitionStyle"].ToString());
                }
            }

            // detach and delete file
            vhdHelper.DetachVirtualHardDisk(vdiPath);
            if (File.Exists(vdiPath))
                File.Delete(vdiPath);
        }

        [TestMethod()]
        public void vhdHelper_partitionWMI_CheckPartitionNumber_NotZero()
        {
            string vdiPath = "C:\\test_part.vhdx";
            char drive = 'Z';
            UInt64 vdiSize = 2;

            // create file
            vhdHelper.CreateFixedVirtualHardDisk(vdiPath, vdiSize);
            // attach file 
            vhdHelper.AttachVirtualHardDisk(vdiPath);

            // Test if there exists partition number with 0
            UInt32 diskNumber = Convert.ToUInt32(vhdHelper.partitionWMI(vdiPath, drive));

            string sqe = string.Format("SELECT * FROM MSFT_Partition WHERE DiskNumber={0}", diskNumber);

            using (ManagementObjectSearcher diskDrives = new ManagementObjectSearcher(@"\\.\Root\Microsoft\Windows\Storage", sqe))
            {
                ManagementObjectCollection moc = diskDrives.Get();
                foreach (ManagementObject diskObj in moc)
                {
                    Assert.AreNotEqual("0", diskObj["PartitionNumber"].ToString());
                }
            }

            // detach and delete the file
            vhdHelper.DetachVirtualHardDisk(vdiPath);
            if (File.Exists(vdiPath))
                File.Delete(vdiPath);
        }

        [TestMethod()]
        public void vhdHelper_formatWMI_CheckFileSystem_IsNTFS()
        {
            string vdiPath = "C:\\test_format.vhdx";
            UInt64 vdiSize = 2;
            char drive = 'Z';
            string driveletter = null;

            // create disk
            vhdHelper.CreateFixedVirtualHardDisk(vdiPath, vdiSize);
            // attach disk 
            vhdHelper.AttachVirtualHardDisk(vdiPath);

            // initialize and partition the disk
            UInt32 diskNumber = Convert.ToUInt32(vhdHelper.partitionWMI(vdiPath, drive));

            // Test if there driveletter has been assigned and check filesystem is NTFS
            vhdHelper.formatWMI(drive.ToString());

            driveletter = drive.ToString() + ":";

            string sqe = string.Format("SELECT * FROM Win32_Volume WHERE DriveLetter='{0}'", driveletter);

            using (ManagementObjectSearcher diskDrives = new ManagementObjectSearcher(@"\root\CIMV2", sqe))
            {
                ManagementObjectCollection moc = diskDrives.Get();
                foreach (ManagementObject diskObj in moc)
                {
                    Assert.AreEqual("NTFS", diskObj["FileSystem"].ToString());
                }
            }

            // detach and delete the disk
            vhdHelper.DetachVirtualHardDisk(vdiPath);
            if (File.Exists(vdiPath))
                File.Delete(vdiPath);
        }

        [TestMethod()]
        public void vhdHelper_defragVHD_CheckDefragmentRecommendation_False()
        {
            string srcPath = "C:\\test";
            string vdiPath = "C:\\test_defrag.vhdx";
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
            // copy files from srcPath to disk
            WinHelper.copyFolderToVhd(srcPath, drive.ToString());
            // defragment VHDX disk
            vhdHelper.defragVHD(drive.ToString());

            // Test if recommendation of defragment of volume is false
            string defragRecommended = null;
            string FilePercentFragmentation = null;

            vhdHelper.defragAnalysis(drive.ToString(), out defragRecommended, out FilePercentFragmentation);

            Assert.AreEqual("False", defragRecommended);

            // detach and delete the file
            vhdHelper.DetachVirtualHardDisk(vdiPath);
            if (File.Exists(vdiPath))
                File.Delete(vdiPath);
        }

        [TestMethod()]
        public void vhdHelper_defragAnalysis_CheckFilePercentFragmentation_NotZero()
        {
            string driveletter = "C";
            string defragRecommended = null;
            string FilePercentFragmentation = null;

            // Test if recommendation of FilePercentFragmentation of volume is not zero
            vhdHelper.defragAnalysis(driveletter, out defragRecommended, out FilePercentFragmentation);

            Assert.AreNotEqual("0", FilePercentFragmentation);
        }
    }
}