using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace HyperVSamples
{
    public class Program
    {
        public static Object _lock = new Object();
        public static Dictionary<string, char> vdipathToDrive;

        public static void test(string srcPath, string vdiPath, UInt64 vdiSize)
        {
            char driveletter = 'Z';
            vdiPath = srcPath + string.Format("_{0}.vhdx", "1");
            vhdHelper.CreateFixedVirtualHardDisk(vdiPath, vdiSize);

            vhdHelper.AttachVirtualHardDisk(vdiPath);

            vhdHelper.partitionWMI(vdiPath, driveletter);

            vhdHelper.formatWMI(driveletter.ToString());

            vhdHelper.defragVHD(driveletter.ToString());
            vhdHelper.DetachVirtualHardDisk(vdiPath);

            WinHelper.verifyVHD(vdiPath, driveletter.ToString());
        }

        static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                // Helper();
                throw new Exception("Need vhdx path and size");
            }

            else
            {
                string srcPath = "C:\\Users\\Administrator\\Desktop\\test";
                string vdiPath = null;
                UInt64 vdiSize = UInt64.Parse("3");

                vdipathToDrive = new Dictionary<string, char>();
                // test(srcPath, vdiPath, vdiSize);


                ParallelOptions threadnum = new ParallelOptions();
                threadnum.MaxDegreeOfParallelism = 4;
                Parallel.For(1, 5, threadnum, i =>
                {
                    lock (_lock)
                    {
                        vdiPath = srcPath + string.Format("_{0}.vhdx", i);
                        createVHDX(srcPath, vdiPath, vdiSize);
                    }
                    manageVHDX(srcPath, vdiPath, vdiSize);
                });
            }
            Console.ReadLine();
        }

        public static void createVHDX(string srcPath, string vdiPath, ulong vdiSize)
        {

            if (vdiPath != null)
            {
                String vdiFolder = Path.GetDirectoryName(vdiPath);
                if (!System.IO.Directory.Exists(vdiFolder))
                {
                    try
                    {
                        System.IO.Directory.CreateDirectory(vdiFolder);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Directory already exists for output vdi path");
                    }
                }
                if (!vdiPath.Contains(".vhdx") && !vdiPath.Contains(".vhd"))
                {
                    Console.WriteLine("Output file should be a .vhd/.vhdx file");
                    throw new Exception();
                }
            }
            else
            {
                Console.WriteLine("Invalid path for output vdi specified");
                Helper();
            }

            //Condition check for srcPath to exist
            if (!Directory.Exists(srcPath))
            {
                Console.WriteLine("Source path for creating vhd doesnt exist");
                throw new Exception("Argument constraint error.");
            }

            Console.WriteLine("Creating empty vhd at '" + vdiPath + "' of size '" + vdiSize + " GB'");

            // Create vhdx
            vhdHelper.CreateFixedVirtualHardDisk(vdiPath, vdiSize);

        }
        public static void manageVHDX(string srcPath, string vdiPath, ulong vdiSize)
        {

            // string driveletter = "";
            UInt32 diskNumber = 0;
            lock (_lock)
            {
                vhdHelper.AttachVirtualHardDisk(vdiPath);

                vdipathToDrive[vdiPath] = WinHelper.getFreeDisks();
                diskNumber = Convert.ToUInt32(vhdHelper.partitionWMI(vdiPath, vdipathToDrive[vdiPath]));

                // Formatting virtual hard disk
                if (vhdHelper.formatWMI(vdipathToDrive[vdiPath].ToString()))
                {
                    Console.WriteLine("{0} Formatting is successful.", vdiPath);
                }
                else
                {
                    Console.WriteLine("{0} Formatting is fail.", vdiPath);
                    throw new Exception();
                }
            }
            // copy folder and content from source path to vhdx drive letter
            Console.WriteLine("Copying contents of '" + srcPath + "' to drive '" + vdipathToDrive[vdiPath].ToString() + "'");
            if (!WinHelper.copyFolderToVhd(srcPath, vdipathToDrive[vdiPath].ToString()))
            {
                Console.WriteLine("Error copying folder data to drive");
                throw new Exception("VDI convert error.");
            }

            // fix permisission of content on the vhdx drive
            Console.WriteLine("Fixing permissions on the created vdi volume at '" + vdipathToDrive[vdiPath].ToString() + "'");
            if (Directory.Exists(vdipathToDrive[vdiPath].ToString() + ":\\asgard"))
            {
                if (!WinHelper.fixPermission(vdipathToDrive[vdiPath].ToString() + ":\\asgard"))
                {
                    Console.WriteLine("Error setting the permissions on the '" + vdipathToDrive[vdiPath].ToString() + ":\\asgard'");
                    throw new Exception("VDI file permission error.");
                }
            }
            if (File.Exists(vdipathToDrive[vdiPath].ToString() + ":\\manifest"))
            {
                if (!WinHelper.fixPermission(vdipathToDrive[vdiPath].ToString() + ":\\manifest"))
                {
                    Console.WriteLine("WARN: Failed setting the permissions on the '" + vdipathToDrive[vdiPath].ToString() + ":\\manifest'");
                }
            }

            Console.WriteLine("{0} Defragmenting the drive '" + vdipathToDrive[vdiPath].ToString() + "' before detaching it", vdiPath);
            if (!vhdHelper.defragVHD(vdipathToDrive[vdiPath].ToString()))
            {
                Console.WriteLine("Error defragmenting the drive '" + vdipathToDrive[vdiPath].ToString() + "'");
                throw new Exception("Defragmenting drive error");
            }

            // detach vhdx
            Console.WriteLine("*******************************************Detaching final vhd at '" + vdiPath + "' from drive '" + vdipathToDrive[vdiPath].ToString() + "'");
            vhdHelper.DetachVirtualHardDisk(vdiPath);

            //lock (_lock)
            //{
            // verify vhdx
            Console.WriteLine("Verifying that the created vhd is good : {0}", vdiPath);
            if (!WinHelper.verifyVHD(vdiPath, vdipathToDrive[vdiPath].ToString()))
            {
                Console.WriteLine("Failed the VHD verification process");
                throw new Exception("VHD goodness error.");
            }
            //}
            
            // delete vhdx
            Console.WriteLine("Deleting the temp vhd file on the local system at '" + vdiPath + "'");
            if (File.Exists(vdiPath))
            {
                try
                {
                    File.Delete(vdiPath);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Exception trying to clean the local vhd at '" + vdiPath + "'");
                }
            }
            
            vdipathToDrive.Remove(vdiPath);
            Console.WriteLine("=============Ending manageVHDX run for the instance===============");
        }

        private static void Helper()
        {
            Console.WriteLine("Wrong arguments specified, usage expected :");
            Console.WriteLine("            vdiconvertor.exe <srcFolder> <path_to_vdi> <size> <share_to_copy> [Exitfile]");
            Console.WriteLine("                       - <srcFolder>  = source folder pointing to which vhich needs to be created ( Just the drive letter alphabet) ");
            Console.WriteLine("                       - path_to_vdi  = Full Path to the output vdi file");
            Console.WriteLine("                       - size  = size of vhd in GB");
            Console.WriteLine("                       - share_to_copy = Share to copy the vdi to in format \\<sharename>");
            throw new Exception("Argument constraint error.");
            //System.Environment.Exit(ARG_CONSTRAINT_ERROR);
        }
    }
}
