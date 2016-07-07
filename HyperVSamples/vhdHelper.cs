using System;
using System.Management;
using HyperVSamplesTest.Utility;
using System.Diagnostics;
using System.IO;

namespace HyperVSamples
{
    public enum VirtualHardDiskType
    {
        Unknown = 0,
        FixedSize = 2,
        DynamicallyExpanding = 3,
        Differencing = 4
    }

    public enum VirtualHardDiskFormat
    {
        Unknown = 0,
        Vhd = 2,
        Vhdx = 3
    }

    public enum VirtualHardDiskPartitionStyle
    {
        MBR = 1,
        GPT = 2
    }

    public enum MbrType
    {
        FAT12 = 1,
        FAT16 = 4,
        Extended = 5,
        Huge = 6,
        IFS = 7,        // An NTFS or ExFAT partition.
        FAT32 = 12
    }

    public class vhdHelper
    {
        public const UInt64 Size1G = 0x40000000;

        public static ManagementObject GetFirstObjectFromCollection(ManagementObjectCollection collection)
        {
            // Get from Common utilities for the virtualization samples (V2) - GetFirstObjectFromCollection
            if (collection.Count == 0)
            {
                throw new ArgumentException("The collection contains no objects", "collection");
            }

            foreach (ManagementObject managementObject in collection)
            {
                return managementObject;
            }

            return null;
        }

        public static ManagementObject GetImageManagementService(ManagementScope scope)
        {
            // Get from Common utilities for the virtualization samples (V2) - GetVirtualMachineManagementService
            using (ManagementClass imageManagementServiceClass = new ManagementClass("Msvm_ImageManagementService"))
            {
                imageManagementServiceClass.Scope = scope;

                ManagementObject imageManagementService = GetFirstObjectFromCollection(imageManagementServiceClass.GetInstances());

                return imageManagementService;
            }
        }

        public static string GetVirtualHardDiskSettingDataEmbeddedInstance(string serverName, string namespacePath, string m_Path, UInt64 m_MaxInternalSize, uint diskformat = (uint)VirtualHardDiskFormat.Vhdx, uint diskType = (uint)VirtualHardDiskType.DynamicallyExpanding)
        {
            ManagementPath path = new ManagementPath()
            {
                Server = serverName,
                NamespacePath = namespacePath,
                ClassName = "Msvm_VirtualHardDiskSettingData"
            };

            using (ManagementClass settingsClass = new ManagementClass(path))
            {
                using (ManagementObject settingsInstance = settingsClass.CreateInstance())
                {
                    settingsInstance["Type"] = diskType;
                    settingsInstance["Format"] = diskformat;
                    settingsInstance["Path"] = m_Path;
                    settingsInstance["MaxInternalSize"] = m_MaxInternalSize * Size1G;
                    //settingsInstance["ParentPath"] = m_ParentPath;
                    //settingsInstance["BlockSize"] = 0;
                    //settingsInstance["LogicalSectorSize"] = 0;
                    //settingsInstance["PhysicalSectorSize"] = 0;

                    string settingsInstanceString = settingsInstance.GetText(TextFormat.WmiDtd20);

                    return settingsInstanceString;
                }
            }
        }

        public static void CreateFixedVirtualHardDisk(string vhdPath, UInt64 maxInternalSize)
        {

            ManagementScope scope = new ManagementScope(@"Root\Virtualization\V2", null);

            using (ManagementObject imageManagementService = GetImageManagementService(scope))
            {
                using (ManagementBaseObject inParams = imageManagementService.GetMethodParameters("CreateVirtualHardDisk"))
                {
                    inParams["VirtualDiskSettingData"] = GetVirtualHardDiskSettingDataEmbeddedInstance("\\", imageManagementService.Path.Path, vhdPath, maxInternalSize);
                    using (ManagementBaseObject outParams = imageManagementService.InvokeMethod("CreateVirtualHardDisk", inParams, null))
                    {
                        WmiUtility.ValidateOutput(outParams, scope, true, false);
                    }
                }
            }
        }

        public static void AttachVirtualHardDisk(string VirtualHardDiskPath)
        {
            ManagementScope scope = new ManagementScope(@"Root\Virtualization\V2", null);

            using (ManagementObject imageManagementService = GetImageManagementService(scope))
            {
                using (ManagementBaseObject inParams = imageManagementService.GetMethodParameters("AttachVirtualHardDisk"))
                {
                    inParams["Path"] = VirtualHardDiskPath;
                    inParams["AssignDriveLetter"] = true;
                    inParams["ReadOnly"] = false;

                    using (ManagementBaseObject outParams = imageManagementService.InvokeMethod("AttachVirtualHardDisk", inParams, null))
                    {
                        // WmiUtility.ValidateOutput(outParams, scope, true, false);
                        if ((UInt32)outParams["ReturnValue"] == ReturnCode.Started)
                        {
                            if (Utility.JobCompleted(outParams, scope))
                            {
                                Console.WriteLine("{0} was attached successfully.", inParams["Path"]);
                            }
                            else
                            {
                                Console.WriteLine("Unable to attach {0}", inParams["Path"]);
                            }
                        }
                    }
                }
            }
        }

        public static void DetachVirtualHardDisk(string path)
        {
            ManagementScope scope = new ManagementScope(@"root\virtualization\V2", null);

            ManagementClass mountedStorageImageServiceClass = new ManagementClass("Msvm_MountedStorageImage");

            mountedStorageImageServiceClass.Scope = scope;

            using (ManagementObjectCollection collection = mountedStorageImageServiceClass.GetInstances())
            {
                foreach (ManagementObject image in collection)
                {
                    using (image)
                    {
                        string name = image.GetPropertyValue("Name").ToString();
                        if (string.Equals(name, path, StringComparison.OrdinalIgnoreCase))
                        {
                            using (ManagementBaseObject outParams = image.InvokeMethod("DetachVirtualHardDisk", null, null))
                            {
                                if ((UInt32)outParams["ReturnValue"] == 0)
                                {
                                    Console.WriteLine("{0} was detached successfully.", path);
                                }
                                else
                                {
                                    Console.WriteLine("Unable to dettach {0}", path);
                                }
                            }
                            break;
                        }
                    }
                }
            }
        }

        public static bool initializeWMI(ManagementObject vhdDisk, uint partitionStyle = (uint)VirtualHardDiskPartitionStyle.MBR)
        {
            using (ManagementBaseObject inParams = vhdDisk.GetMethodParameters("Initialize"))
            {
                inParams["PartitionStyle"] = partitionStyle;

                using (ManagementBaseObject outParams = vhdDisk.InvokeMethod("Initialize", inParams, null))
                {
                    if ((UInt32)outParams["ReturnValue"] == 0)
                        return true;
                    else
                        return false;
                }
            }
        }

        public static string partitionWMI(string vdiPath, char driveletter)
        {
            try
            {
                var scope = new ManagementScope(@"\\.\Root\Microsoft\Windows\Storage");
                scope.Connect();

                const string query = "SELECT * FROM MSFT_Disk";
                var objectQuery = new ObjectQuery(query);
                var seacher = new ManagementObjectSearcher(scope, objectQuery);
                var disks = seacher.Get();
                string diskNumber = null;
                ManagementObject disk = null;

                foreach (ManagementObject diskObj in disks)
                {
                    //Console.WriteLine("All Number: {0}", diskObj["Number"]);
                    //Console.WriteLine("All NumberOfPartitions: {0}", diskObj["NumberOfPartitions"]);
                    //Console.WriteLine("All SerialNumber: {0}", diskObj["SerialNumber"]);
                    //Console.WriteLine("All UniqueId: {0}", diskObj["UniqueId"]);
                    //Console.WriteLine("All Signature: {0}", diskObj["Signature"]);

                    if (diskObj["NumberOfPartitions"].Equals((uint)0) && diskNumber == null && disk == null)
                    {
                        Console.WriteLine("Number: {0} -- {1}", diskObj["Number"], vdiPath);
                        Console.WriteLine("NumberOfPartitions: {0} -- {1}", diskObj["NumberOfPartitions"], vdiPath);
                        disk = diskObj;
                        diskNumber = diskObj["Number"].ToString();
                    }
                }

                if (diskNumber == null || disk == null)
                {
                    Console.WriteLine("\nPartition is fail. --- {0}\n", vdiPath);
                    return null;
                }

                /// Initialize the virtual disk
                if (initializeWMI(disk, (uint)VirtualHardDiskPartitionStyle.MBR))
                {
                    Console.WriteLine("\nInitialize is successful. --- {0} \n", vdiPath);
                }
                else
                {
                    Console.WriteLine("\nInitialize is fail. --- {0} \n", vdiPath);
                    return null;
                }

                //var disk = disks.Cast<ManagementObject>().FirstOrDefault();

                var parameters = disk.GetMethodParameters("CreatePartition");

                parameters["UseMaximumSize"] = true;
                parameters["MbrType"] = MbrType.IFS;
                // If TRUE, the next available drive letter will be assigned to the created partition. 
                // If no more drive letters are available, the partition will be created with no drive letter. 
                // This parameter cannot be used with DriveLetter. If both parameters are specified, an Invalid Parameter error will be returned.
                // parameters["AssignDriveLetter"] = true;
                parameters["DriveLetter"] = driveletter;
                parameters["IsActive"] = true;

                var result = disk.InvokeMethod("CreatePartition", parameters, null);
                Console.WriteLine("ReturnValue. ---  {0}\n", result["ReturnValue"]);
                if ((UInt32)result["ReturnValue"] == 0)
                {
                    Console.WriteLine("Partition is successful. ---  {0}\n", vdiPath);
                    return diskNumber;
                }
                else
                    return null;
            }
            catch (Exception exception)
            {
                Debug.Fail(exception.Message);
                return null;
            }
        }

        public static bool formatWMI(string drive, string fileSystem = "NTFS", bool quickformat = true, string label = "")
        {
            
            if (drive == null || drive == "C")
                return false;

            string driveLetter = drive+":";
            try
            {
                string sq = string.Format("SELECT * FROM Win32_Volume WHERE DriveLetter='{0}'", driveLetter);
                //string sq = string.Format("SELECT * FROM MSFT_Volume WHERE DriveLetter='{0}'", drive);
                ManagementObject vhdDisk = null;
                using (ManagementObjectSearcher diskDrives = new ManagementObjectSearcher(@"\root\CIMV2", sq))
                // using (ManagementObjectSearcher diskDrives = new ManagementObjectSearcher(@"\Root\Microsoft\Windows\Storage", sq))
                {
                    ManagementObjectCollection moc = diskDrives.Get();
                    foreach (ManagementObject diskObj in moc)
                    {
                        Console.WriteLine("Win32_Volume DriveLetter: {0}", diskObj["DriveLetter"]);
                        // Console.WriteLine("Win32_Volume Path: {0}", diskObj["Path"]);
                        vhdDisk = diskObj;
                    }

                    if (vhdDisk == null)
                        return false;

                    using (ManagementBaseObject inParams = vhdDisk.GetMethodParameters("Format"))
                    {
                        inParams["FileSystem"] = fileSystem;
                        inParams["QuickFormat"] = quickformat;
                        inParams["Label"] = label;

                        //inParams["FileSystem"] = fileSystem;
                        //inParams["FileSystemLabel"] = label;
                        //inParams["AllocationUnitSize"] = 4096;
                        //inParams["Full"] = false;
                        //inParams["Force"] = false;
                        //inParams["Compress"] = false;
                        //inParams["ShortFileNameSupport"] = false;
                        //inParams["SetIntegrityStreams"] = false;
                        //inParams["UseLargeFRS"] = false;
                        //inParams["DisableHeatGathering"] = false;


                        using (ManagementBaseObject outParams = vhdDisk.InvokeMethod("Format", inParams, null))
                        {
                            Console.WriteLine("Format Win32_Volume ReturnValue: {0}", outParams["ReturnValue"]);
                            if ((UInt32)outParams["ReturnValue"] == 0)
                                return true;
                            else
                                return false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

            return true;
        }
        /*
        public static string assignDriveLetter(string uniqueId)
        {
            
            UInt32 diskNumber = 0;
            //string sqery_disk = string.Format("SELECT * FROM MSFT_Disk WHERE UniqueId='{0}'", uniqueId);
            string sqery_disk = string.Format("SELECT * FROM MSFT_Disk");
            using (ManagementObjectSearcher diskDrives = new ManagementObjectSearcher(@"\\.\Root\Microsoft\Windows\Storage", sqery_disk))
            {
                ManagementObjectCollection moc = diskDrives.Get();
                foreach (ManagementObject diskObj in moc)
                {

                    // Console.WriteLine("assignDriveLetter All Number: {0}", diskObj["Number"]);
                    //Console.WriteLine("assignDriveLetter NumberOfPartitions: {0}", diskObj["NumberOfPartitions"]);
                    //Console.WriteLine("assignDriveLetter SerialNumber: {0}", diskObj["SerialNumber"]);
                    // Console.WriteLine("assignDriveLetter All UniqueId: {0}", diskObj["UniqueId"]);
                    //Console.WriteLine("assignDriveLetter Signature: {0}", diskObj["Signature"]);
                    if (diskObj["UniqueId"].Equals(uniqueId))
                    {
                        diskNumber = Convert.ToUInt32(diskObj["Number"].ToString());
                        Console.WriteLine("assignDriveLetter selected UniqueId: {0}", diskObj["UniqueId"]);
                        Console.WriteLine("assignDriveLetter selected Number: {0}", diskObj["Number"]);
                    }
                }
            }
            Console.WriteLine("------------------------ assignDriveLetter Disk Number : {0} -----------------------------------", diskNumber);

            if (diskNumber == 0)
            {
                Console.WriteLine("Getting the wrong disk number :{0}", diskNumber);
                throw new Exception();
            }

            string sqe = string.Format("SELECT * FROM MSFT_Partition WHERE DiskNumber={0}", diskNumber);
            //string sqe = string.Format("SELECT * FROM MSFT_Disk WHERE Number={0}", diskNumber);
            ManagementObject vhdDisk = null;

            using (ManagementObjectSearcher diskDrives = new ManagementObjectSearcher(@"\\.\Root\Microsoft\Windows\Storage", sqe))
            {
                ManagementObjectCollection moc = diskDrives.Get();
                foreach (ManagementObject diskObj in moc)
                {
                    vhdDisk = diskObj;
                }

                if (vhdDisk == null)
                {
                    Console.WriteLine("assign DriveLetter is failed : vhdDisk == null");
                    throw new Exception();
                }

                using (ManagementBaseObject inParams = vhdDisk.GetMethodParameters("AddAccessPath"))
                {
                    inParams["AssignDriveLetter"] = true;

                    using (ManagementBaseObject outParams = vhdDisk.InvokeMethod("AddAccessPath", inParams, null))
                    {
                        if ((UInt32)outParams["ReturnValue"] == 0)
                            Console.WriteLine("assign DriveLetter is successful");
                        else
                        {
                            Console.WriteLine("assign DriveLetter is success but unknown return value: {0}", outParams["ReturnValue"]);
                        }
                    }
                }
            }

            // retrieve drive letter
            string driveLetter = null;
            string sqee = string.Format("SELECT * FROM MSFT_Partition WHERE DiskNumber={0}", diskNumber);
            string drive = "";

            using (ManagementObjectSearcher diskDrives = new ManagementObjectSearcher(@"\\.\Root\Microsoft\Windows\Storage", sqee))
            {
                ManagementObjectCollection moc = diskDrives.Get();
                foreach (ManagementObject diskObj in moc)
                {

                    Console.WriteLine("assignDriveLetter  DiskNumber: {0}", diskObj["DiskNumber"]);
                    Console.WriteLine("assignDriveLetter  MSFT_Partition DriveLetter: {0}", diskObj["DriveLetter"]);

                    driveLetter = diskObj["DriveLetter"].ToString() + ":";
                    drive = diskObj["DriveLetter"].ToString();

                }
            }

            if (driveLetter == null || driveLetter == "C:" || driveLetter == "D:")
                return "";

            return drive;
        }
        */
        
        public static bool defragVHD(string drive)
        {
            try
            {
                
                if (drive == null || drive == "C" )
                    throw new Exception();

                string sq = string.Format("SELECT * FROM MSFT_Volume WHERE DriveLetter='{0}'", drive);
                ManagementObject vhdDisk = null;
                using (ManagementObjectSearcher diskDrives = new ManagementObjectSearcher(@"\Root\Microsoft\Windows\Storage", sq))
                {
                    ManagementObjectCollection moc = diskDrives.Get();
                    foreach (ManagementObject diskObj in moc)
                    {
                        vhdDisk = diskObj;
                        // Console.WriteLine("Defrag - DriveLetter :{0}", diskObj["DriveLetter"]);
                    }

                    if (vhdDisk == null)
                        return false;

                    using (ManagementBaseObject inParams = vhdDisk.GetMethodParameters("Optimize"))
                    {
                        inParams["ReTrim"] = true;
                        inParams["Analyze"] = true;
                        inParams["Defrag"] = true;
                        inParams["SlabConsolidate"] = true;
                        inParams["TierOptimize"] = true;

                        using (ManagementBaseObject outParams = vhdDisk.InvokeMethod("Optimize", inParams, null))
                        {
                            Console.WriteLine("+++++++++++++++++++++++++++++ outParams - ReturnValue : {0}", outParams["ReturnValue"]);
                            if ((UInt32)outParams["ReturnValue"] == 0)
                            {
                                Console.WriteLine("Degragment disk is successful.");
                                return true;
                            }
                            else
                                return false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return true;
        }

        public static void defragAnalysis(string drive, out string defragRecommended, out string FilePercentFragmentation)
        {

            string driveletter = drive + ":";
            string sqe = string.Format("SELECT * FROM Win32_Volume WHERE DriveLetter='{0}'", driveletter);
            ManagementObject disk = null;
            using (ManagementObjectSearcher diskDrives = new ManagementObjectSearcher(@"\root\CIMV2", sqe))
            {
                ManagementObjectCollection moc = diskDrives.Get();
                foreach (ManagementObject diskObj in moc)
                {
                    Console.WriteLine("Win32_Volume DriveLetter: {0}", diskObj["DriveLetter"]);

                    disk = diskObj;
                }

                if (disk == null)
                    throw new Exception();


                using (ManagementBaseObject inParams = disk.GetMethodParameters("DefragAnalysis"))
                {
                    using (ManagementBaseObject outParams = disk.InvokeMethod("DefragAnalysis", inParams, null))
                    {
                        defragRecommended = outParams["DefragRecommended"].ToString();

                        ManagementBaseObject aa = outParams["DefragAnalysis"] as ManagementBaseObject;

                        Console.WriteLine("defragRecommended : {0}", defragRecommended);
                        Console.WriteLine("aa : {0}", aa);
                        Console.WriteLine("FilePercentFragmentation : {0}", aa["FilePercentFragmentation"]);
                        Console.WriteLine("FragmentedFolders : {0}", aa["FragmentedFolders"]);
                        Console.WriteLine("TotalFragmentedFiles : {0}", aa["TotalFragmentedFiles"]);
                        Console.WriteLine("AverageFragmentsPerFile : {0}", aa["AverageFragmentsPerFile"]);

                        Console.WriteLine("ReturnValue : {0}", outParams["ReturnValue"]);

                        FilePercentFragmentation = aa["FilePercentFragmentation"].ToString();
                    }
                }
            }

        }

    }
}
