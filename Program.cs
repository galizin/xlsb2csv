using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.IO.Compression;

namespace xlsbtocsv
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                using (FileStream zipToOpen = new FileStream(@"C:\user_main\python\FL_insurance_sample3.xlsb", FileMode.Open, FileAccess.Read))
                {
                    using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Read))
                    {
                        Dictionary<uint, string> shstr = new Dictionary<uint, string>();
                        List<uint> datestyle = new List<uint>();
                        using (Stream ms = archive.GetEntry(@"xl/styles.bin").Open())
                        {
                            loadstyles(ms, ref datestyle);
                        }
                        using (Stream ms = archive.GetEntry(@"xl/sharedStrings.bin").Open())
                        {
                            loadsharedstrings(ms, ref shstr);
                        }
                        using (Stream ms = archive.GetEntry(@"xl/worksheets/sheet1.bin").Open())
                        {
                            readworksheet(ms, shstr, datestyle);
                        }
                    }
                }
            }
            catch (FileNotFoundException ioEx)
            {
                Console.WriteLine(ioEx.Message);
            }
        }
        static public void loadstyles(Stream fsSource, ref List<uint> outDateStyles)
        {
            List<ushort> datestyles = new List<ushort>();
            for (int i = 14; i < 23; i++)
                datestyles.Add((ushort)i);
            for (int i = 45; i < 48; i++)
                datestyles.Add((ushort)i);
            uint styleid = 0;
            Dictionary<uint, ushort> xf = new Dictionary<uint, ushort>();
            while (1 == 1)
            {
                int rec_id;
                byte[] data = null;
                readrecord(out rec_id, ref data, fsSource);
                if (rec_id == -1)
                    break;
                switch (rec_id)
                {
                    case 44: // custom
                        string filteredstring = getxlwidestring(data, 2).Replace("[Black]", "").Replace("[Green]", "").Replace("[White]", "").Replace("[Blue]", "").Replace("[Magenta]", "").Replace("[Yellow]", "").Replace("[Cyan]", "").Replace("[Red]", "").ToLower();
                        if (filteredstring.Contains("y") || filteredstring.Contains("d") || filteredstring.Contains("h") || filteredstring.Contains("m") || filteredstring.Contains("s") || filteredstring.Contains("a") || filteredstring.Contains("p"))
                        {
                            datestyles.Add(BitConverter.ToUInt16(data, 0));
                        }
                        break;
                    case 47: // BrtXF
                        xf.Add(styleid, BitConverter.ToUInt16(data, 2));
                        styleid++;
                        break;
                }
            }
            for (uint i = 0; i < styleid; i++)
            {
                if (datestyles.Contains(xf[i]))
                    outDateStyles.Add(i);
            }
        }
        static public void loadsharedstrings(Stream fsSource, ref Dictionary<uint, string> shstr)
        {
            uint strid = 0;
            while (1 == 1)
            {
                int rec_id;
                byte[] data = null;
                readrecord(out rec_id, ref data, fsSource);
                if (rec_id == -1)
                    break;
                switch (rec_id)
                {
                    case 19: // Shared string
                        shstr.Add(strid, getxlwidestring(data, 1));
                        strid++;
                        break;
                }
            }
        }
        static public void readworksheet(Stream fsSource, Dictionary<uint, string> shstr, List<uint> datestyles)
        {
            using (StreamWriter outputFile = new StreamWriter("output.txt", false, Encoding.UTF8))
            {
                while (1 == 1)
                {
                    int rec_id;
                    byte[] data = null;
                    readrecord(out rec_id, ref data, fsSource);
                    if (rec_id == -1)
                        break;
                    switch (rec_id)
                    {
                        case 0: // row
                            outputFile.WriteLine("row {0}", BitConverter.ToUInt32(data, 0));
                            break;
                        case 1: // BrtCellBlank
                            writecellinfo(outputFile, data, datestyles);
                            outputFile.WriteLine("blank cell");
                            break;
                        case 2: // BrtCellRk
                            writecellinfo(outputFile, data, datestyles);
                            uint value = BitConverter.ToUInt32(data, 8);
                            double x;
                            bool div100 = (data[8] & 1u) == 1u;
                            bool fltype = (data[8] & 2u) == 0u;
                            if (fltype)
                            {
                                byte[] dbl = new byte[8];
                                dbl[0] = 0;
                                dbl[1] = 0;
                                dbl[2] = 0;
                                dbl[3] = 0;
                                dbl[4] = (byte)(data[8] & 0xFC);
                                dbl[5] = data[9];
                                dbl[6] = data[10];
                                dbl[7] = data[11];
                                x = BitConverter.ToDouble(dbl, 0);
                            }
                            else
                            {
                                x = Convert.ToDouble(value >> 2);
                            }
                            if (div100)
                                x = x / 100;
                            if (dateformatted(data, datestyles))
                                outputFile.WriteLine("rk date {0}", stringdate(x));
                            else
                                outputFile.WriteLine("rk {0}", x);
                            break;
                        case 3: // BrtCellError
                            writecellinfo(outputFile, data, datestyles);
                            outputFile.WriteLine("err {0}", data[8]);
                            break;
                        case 4: // BrtCellBool
                            writecellinfo(outputFile, data, datestyles);
                            outputFile.WriteLine("bool {0}", data[8]);
                            break;
                        case 5: // BrtCellReal
                            writecellinfo(outputFile, data, datestyles);
                            if (dateformatted(data, datestyles))
                                outputFile.WriteLine("dbl date {0}", stringdate(BitConverter.ToDouble(data, 8)));
                            else
                                outputFile.WriteLine("dbl {0}", BitConverter.ToDouble(data, 8));
                            break;
                        case 6: // BrtCellSt
                            writecellinfo(outputFile, data, datestyles);
                            outputFile.WriteLine("value {0}", getxlwidestring(data, 8));
                            break;
                        case 7: // BrtCellIsst
                            writecellinfo(outputFile, data, datestyles);
                            outputFile.WriteLine("isst {0}", shstr[BitConverter.ToUInt32(data, 8)]);
                            break;
                        case 8: // BrtFmlaString
                            break;
                        case 9: // BrtFmlaNum 
                            break;
                        case 10: // BrtFmlaBool 
                            break;
                        case 11: // BrtFmlaError
                            break;
                        case 19: // Shared string
                            outputFile.WriteLine("shstr {0}", fsSource.Position);
                            break;
                        case 44: // fmt
                            outputFile.WriteLine("fmt {0}", fsSource.Position);
                            break;
                        default:
                            break;
                    }
                }
            }
        }
        static public bool dateformatted(byte[] data, List<uint> datastyles)
        {
            uint styleid;
            getcellno(data, out styleid);
            return datastyles.Contains(styleid);
        }
        static public void writecellinfo(StreamWriter f, byte[] data, List<uint> datastyles)
        {
            uint styleid;
            f.Write("col {0} style {1} ", getcellno(data, out styleid), styleid);
        }
        static public uint getcellno(byte[] buffer, out uint styleid)
        {
            styleid = buffer[4] + buffer[5] * 256u + buffer[6] * 256u * 256u;
            return BitConverter.ToUInt32(buffer, 0);
        }
        static public string getxlwidestring(byte[] buffer, int pos)
        {
            int strlen = Convert.ToInt32(BitConverter.ToUInt32(buffer, pos));
            return System.Text.Encoding.Unicode.GetString(buffer, pos + 4, strlen * 2);
        }
        static public void readrecord(out int rec_id, ref byte[] data, Stream fsSource)
        {
            rec_id = read_id(fsSource);
            if (rec_id == -1)
                return;
            int rec_len = read_len(fsSource);
            data = new byte[rec_len];
            fsSource.Read(data, 0, rec_len);
        }
        static public int read_id(Stream fsSource)
        {
            int b = fsSource.ReadByte();
            if (b == -1)
                return -1;
            if (b < 128)
                return b;
            else
            {
                int b2 = fsSource.ReadByte();
                if (b2 == -1)
                    return -1;
                return b2 * 128 + (b - 128);
            }
        }
        static public int read_len(Stream fsSource)
        {
            int multiplier = 1;
            int accumulated = 0;
            for (int i = 0; i < 4; i++)
            {
                int b = fsSource.ReadByte();
                if (b == -1)
                    return -1;
                if (i == 3 && b > 127)
                    b -= 128;
                if (b < 128)
                    return b * multiplier + accumulated;
                accumulated += b * multiplier;
                multiplier = multiplier * 128;
            }
            throw new IndexOutOfRangeException("unable to calculate record length");
        }
        static public string stringdate(double innumeric)
        {
            if (Math.Truncate(innumeric) == 0)
            {
                return new DateTime(1900, 1, 1, 0, 0, 0).AddSeconds(Math.Truncate(innumeric * 24 * 60 * 60)).ToString("o");
            }
            else
            {
                if (Math.Truncate(innumeric) >= 61)
                    innumeric -= 1;
                // According to Lotus 1-2-3, Feb 29th 1900 is a real thing, therefore we have to remove one day after that date
                return new DateTime(1899, 12, 31, 0, 0, 0).AddDays(Math.Truncate(innumeric)).AddSeconds(Math.Truncate((innumeric % 1) * 24 * 60 * 60)).ToString("o");
                //  else
                // Feb 29th 1900 will show up as Mar 1st 1900 because Python won't handle that date
                //  return new DateTime(1899, 12, 31, 0, 0, 0) + timedelta(days=int(date), seconds=int((date % 1) * 24 * 60 * 60));
            }
        }
    }
}

