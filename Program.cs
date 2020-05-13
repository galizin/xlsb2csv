using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace xlsbtocsv
{
    internal class Program
    {
        private const string stylefname = @"xl/styles.bin";
        private const string stringfname = @"xl/sharedStrings.bin";
        private const string sheetfname = @"xl/worksheets/sheet1.bin";
        private static void Main()
        {
            try
            {
                string filename = Directory.GetFiles(Environment.CurrentDirectory).FirstOrDefault(a => a.EndsWith(".xlsb"));
                using (FileStream zipToOpen = new FileStream(filename, FileMode.Open, FileAccess.Read))
                {
                    using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Read))
                    {
                        Dictionary<uint, string> shstr = new Dictionary<uint, string>();
                        List<uint> datestyle = new List<uint>();
                        if (archive.Entries.FirstOrDefault(a => a.FullName == stylefname) != null)
                            using (Stream ms = archive.GetEntry(stylefname).Open())
                            {
                                Loadstyles(ms, ref datestyle);
                            }
                        if (archive.Entries.FirstOrDefault(a => a.FullName == stringfname) != null)
                            using (Stream ms = archive.GetEntry(stringfname).Open())
                            {
                                Loadsharedstrings(ms, ref shstr);
                            }
                        if (archive.Entries.FirstOrDefault(a => a.FullName == sheetfname) != null)
                            using (Stream ms = archive.GetEntry(sheetfname).Open())
                            {
                                Readworksheet(ms, shstr, datestyle, filename);
                            }
                    }
                }
            }
            catch (FileNotFoundException ioEx)
            {
                Console.WriteLine(ioEx.Message);
            }
        }
        public static void Loadstyles(Stream fsSource, ref List<uint> outDateStyles)
        {
            List<ushort> datestyles = new List<ushort>();
            for (int i = 14; i < 23; i++)
            {
                datestyles.Add((ushort)i);
            }
            for (int i = 45; i < 48; i++)
            {
                datestyles.Add((ushort)i);
            }
            uint styleid = 0;
            Dictionary<uint, ushort> xf = new Dictionary<uint, ushort>();
            //using (StreamWriter outputFile = new StreamWriter("styleinfo.txt", false, Encoding.UTF8))
            {
                bool cellxfsection = false;
                while (1 == 1)
                {
                    byte[] data = null;
                    Readrecord(out int rec_id, ref data, fsSource);
                    if (rec_id == -1)
                    {
                        break;
                    }
                    //outputFile.WriteLine("{0} {1}", rec_id, BitConverter.ToString(data));
                    switch (rec_id)
                    {
                        case 44: // custom
                            string filteredstring = Getxlwidestring(data, 2).Replace("[Black]", "").Replace("[Green]", "").Replace("[White]", "").Replace("[Blue]", "").Replace("[Magenta]", "").Replace("[Yellow]", "").Replace("[Cyan]", "").Replace("[Red]", "").ToLower();
                            if (filteredstring.Contains("y") || filteredstring.Contains("d") || filteredstring.Contains("h") || filteredstring.Contains("m") || filteredstring.Contains("s") || filteredstring.Contains("a") || filteredstring.Contains("p"))
                            {
                                datestyles.Add(BitConverter.ToUInt16(data, 0));
                            }
                            break;
                        case 47: // BrtXF
                            if (cellxfsection)
                            {
                                xf.Add(styleid, BitConverter.ToUInt16(data, 2));
                                styleid++;
                            }
                            break;
                        case 617:
                            cellxfsection = true;
                            break;
                        case 618:
                            cellxfsection = false;
                            break;
                    }
                }
            }
            for (uint i = 0; i < styleid; i++)
            {
                if (datestyles.Contains(xf[i]))
                {
                    outDateStyles.Add(i);
                }
            }
        }
        public static void Loadsharedstrings(Stream fsSource, ref Dictionary<uint, string> shstr)
        {
            //using (StreamWriter outputFiledbg = new StreamWriter("shstr.txt", false, Encoding.UTF8))
            //{
            //using (StreamWriter outputFiledbg1 = new StreamWriter("shstrstruct.txt", false, Encoding.UTF8))
            //{
            uint strid = 0;
            while (1 == 1)
            {
                byte[] data = null;
                Readrecord(out int rec_id, ref data, fsSource);
                if (rec_id == -1)
                {
                    break;
                }
                //outputFiledbg1.WriteLine("{0} {1}", rec_id, BitConverter.ToString(data));
                switch (rec_id)
                {
                    case 19: // Shared string
                        shstr.Add(strid, Getxlwidestring(data, 1));
                        //outputFiledbg.WriteLine("{0} {1}", strid, Getxlwidestring(data, 1));
                        strid++;
                        break;
                    default:

                        break;
                }
            }
            //}
            //}
        }
        public static void Readworksheet(Stream fsSource, Dictionary<uint, string> shstr, List<uint> datestyles, string fname)
        {
            //using (StreamWriter outputFiledbg = new StreamWriter(Path.GetFileNameWithoutExtension(fname) + "shstr.txt", false, Encoding.UTF8))
            //{

            using (StreamWriter outputFile = new StreamWriter(Path.GetFileNameWithoutExtension(fname) + ".txt", false, Encoding.UTF8))
            {
                Console.WriteLine("converting {0}", fname);
                Console.Write("reached row ");
                int rowno = 0;
                int lastmsglen = 0;
                uint lastcol = 0;
                bool firstline = true;
                while (1 == 1)
                {
                    byte[] data = null;
                    Readrecord(out int rec_id, ref data, fsSource);
                    if (rec_id == -1)
                    {
                        break;
                    }
                    switch (rec_id)
                    {
                        case 0: // row
                            if (firstline)
                            {
                                firstline = false;
                                lastcol = 0;
                            }
                            else
                            {
                                outputFile.Write("\r\n");
                                rowno++;
                            }
                            if (rowno % 10000 == 0)
                            {
                                if (lastmsglen != 0)
                                    Console.Write(new String('\b', lastmsglen));
                                Console.Write(rowno.ToString());
                                lastmsglen = rowno.ToString().Length;
                            }
                            break;
                        case 1: // BrtCellBlank
                            WriteCellSeparator(outputFile, data, datestyles, ref lastcol);
                            //outputFile.WriteLine("blank cell");
                            break;
                        case 2: // BrtCellRk
                            WriteCellSeparator(outputFile, data, datestyles, ref lastcol);
                            //outputFile.Write("rk " + BitConverter.ToString(data) + " "); //debug
                            //int value = BitConverter.ToInt32(data, 8);
                            double x;
                            bool div100 = (data[8] & 1u) == 1u;
                            bool fltype = (data[8] & 2u) == 0u;
                            if (fltype)
                            {
                                //outputFile.Write(" type 0 ");
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
                                //int value2 = value >> 2;
                                //int i = -4000137;
                                //i = Convert.ToInt32(value2);
                                x = BitConverter.ToInt32(data, 8) >> 2;//value >> 2;
                                //outputFile.WriteLine(" type 1 {0} {1} {2} {3}", value, value >> 2, (int)(data[8] >> 2) + (data[9] << 6) + (data[10] << 14) + (data[11] << 22), x);
                                //outputFile.WriteLine(Convert.ToString(value, 2) + " value " + value.ToString());
                                //outputFile.WriteLine(Convert.ToString(value2, 2) + " value >> 2 " + value2.ToString());
                                //outputFile.WriteLine(Convert.ToString(i, 2) + " i " + i.ToString());
                                //outputFile.WriteLine(Convert.ToString(-1000035, 2) + " -1000035 ");
                                //outputFile.WriteLine(Convert.ToString((int)(data[8] >> 2) + (data[9] << 6) + (data[10] << 14) + (data[11] << 22), 2) + " (int) data >> ");
                                //outputFile.Write(Convert.ToString((data[8] >> 2) + (data[9] << 6) + (data[10] << 14) + (data[11] << 22), 2) + " data >> ");
                                //outputFile.Write(Convert.ToString(x, 2) + " x ");
                            }
                            if (div100)
                            {
                                //outputFile.Write(" div 100 ");
                                x /= 100;
                            }
                            if (Dateformatted(data, datestyles))
                            {
                                //outputFile.WriteLine("rk date {0}", Stringdate(x));
                                outputFile.Write("{0}", Stringdate(x));
                            }
                            else
                            {
                                //outputFile.WriteLine("rk {0}", x);
                                outputFile.Write("{0}", x);
                            }
                            break;
                        case 3: // BrtCellError
                        case 11: // BrtFmlaError
                            WriteCellSeparator(outputFile, data, datestyles, ref lastcol);
                            string err = "";
                            switch (data[8])
                            {
                                case (byte)0x00u:
                                    err = "#NULL!";
                                    break;
                                case (byte)0x07u:
                                    err = "#DIV/0!";
                                    break;
                                case (byte)0x0Fu:
                                    err = "#VALUE!";
                                    break;
                                case (byte)0x17u:
                                    err = "#REF!";
                                    break;
                                case (byte)0x1Du:
                                    err = "#NAME?";
                                    break;
                                case (byte)0x24u:
                                    err = "#NUM!";
                                    break;
                                case (byte)0x2Au:
                                    err = "#N/A";
                                    break;
                                case (byte)0x2Bu:
                                    err = "#GETTING_DATA";
                                    break;
                                default:
                                    break;
                            }
                            outputFile.Write(err);//data[8]);
                            break;
                        case 4: // BrtCellBool
                        case 10: // BrtFmlaBool 
                            WriteCellSeparator(outputFile, data, datestyles, ref lastcol);
                            outputFile.Write("{0}", data[8]);
                            break;
                        case 5: // BrtCellReal
                        case 9: // BrtFmlaNum 
                            WriteCellSeparator(outputFile, data, datestyles, ref lastcol);
                            if (Dateformatted(data, datestyles))
                            {
                                outputFile.Write("{0}", Stringdate(BitConverter.ToDouble(data, 8)));
                            }
                            else
                            {
                                outputFile.Write("{0}", BitConverter.ToDouble(data, 8));
                            }
                            break;
                        case 6: // BrtCellSt 
                        case 8: //BrtFmlaString
                            WriteCellSeparator(outputFile, data, datestyles, ref lastcol);
                            outputFile.Write("{0}", StringToCsv(Getxlwidestring(data, 8)));
                            break;
                        case 7: // BrtCellIsst
                            WriteCellSeparator(outputFile, data, datestyles, ref lastcol);
                            //if (!shstr.ContainsKey(BitConverter.ToUInt32(data, 8)))
                            //outputFiledbg.WriteLine("shstr {0} missing", BitConverter.ToUInt32(data, 8));
                            //else
                            outputFile.Write("{0}", StringToCsv(shstr[BitConverter.ToUInt32(data, 8)]));
                            break;
                        //case 19: // Shared string
                        //    outputFiledbg.WriteLine("shstr {0}", Getxlwidestring(data, 1));
                        //    shstr.Add();
                        //    break;
                        //case 44: // fmt
                        //    outputFile.WriteLine("fmt {0}", fsSource.Position);
                        //    break;
                        default:
                            break;
                    }
                }
                if (lastmsglen != 0)
                    Console.Write(new String('\b', lastmsglen));
                Console.WriteLine(rowno.ToString());
                Console.WriteLine("finished");
            }
            //}
        }
        public static bool Dateformatted(byte[] data, List<uint> datastyles)
        {
            Getcellno(data, out uint styleid);
            return datastyles.Contains(styleid);
        }
        public static void WriteCellSeparator(StreamWriter f, byte[] data, List<uint> datastyles, ref uint lastcolno)
        {
            //f.Write("col {0} style {1} ", getcellno(data, out uint styleid), styleid);
            int addtabs = 0;
            f.Write(IsFirstCell(data, ref lastcolno, ref addtabs) ? "" : "\t");
            if (addtabs != 0)
                f.Write(new String('\t', addtabs));
        }
        public static bool IsFirstCell(byte[] data, ref uint lastcolno, ref int addtabs)
        {
            uint cellno = Getcellno(data, out _);
            if (cellno != 0u)
            {
                if (cellno != lastcolno + 1u)
                    addtabs = (int)(cellno - lastcolno - 1u);
            }
            lastcolno = cellno;
            return (cellno == 0u);
        }
        public static uint Getcellno(byte[] buffer, out uint styleid)
        {
            styleid = buffer[4] + buffer[5] * 256u + buffer[6] * 256u * 256u;
            return BitConverter.ToUInt32(buffer, 0);
        }
        public static string Getxlwidestring(byte[] buffer, int pos)
        {
            int strlen = Convert.ToInt32(BitConverter.ToUInt32(buffer, pos));
            return System.Text.Encoding.Unicode.GetString(buffer, pos + 4, strlen * 2);
        }
        public static void Readrecord(out int rec_id, ref byte[] data, Stream fsSource)
        {
            rec_id = Read_id(fsSource);
            if (rec_id == -1)
            {
                return;
            }
            int rec_len = Read_len(fsSource);
            data = new byte[rec_len];
            fsSource.Read(data, 0, rec_len);
        }
        public static int Read_id(Stream fsSource)
        {
            int b = fsSource.ReadByte();
            if (b == -1)
            {
                return -1;
            }

            if (b < 128)
            {
                return b;
            }
            else
            {
                int b2 = fsSource.ReadByte();
                if (b2 == -1)
                {
                    return -1;
                }

                return b2 * 128 + (b - 128);
            }
        }
        public static int Read_len(Stream fsSource)
        {
            //int multiplier = 1;
            uint accumulated = 0;
            for (int i = 0; i < 4; i++)
            {
                int b = fsSource.ReadByte();
                if (b == -1)
                {
                    return -1;
                }

                uint currval = (byte)b;
                accumulated += (currval & 0x7F) << (7 * i);
                if ((currval & 0x80) == 0u)
                    return (int)accumulated;
                //if (i == 3 && b > 127)
                //{
                //    b -= 128;
                //}

                //if (b < 128)
                //{
                //    return b * multiplier + accumulated;
                //}

                //accumulated += b * multiplier;
                //multiplier *= 128;


            }
            throw new IndexOutOfRangeException("unable to calculate record length");
        }
        public static string Stringdate(double innumeric)
        {
            if (Math.Truncate(innumeric) == 0)
            {
                return new DateTime(1900, 1, 1, 0, 0, 0).AddSeconds(Math.Truncate(innumeric * 24 * 60 * 60)).ToString("o");
            }
            else
            {
                if (Math.Truncate(innumeric) >= 61)
                {
                    innumeric -= 1;
                }
                // According to Lotus 1-2-3, Feb 29th 1900 is a real thing, therefore we have to remove one day after that date
                return new DateTime(1899, 12, 31, 0, 0, 0).AddDays(Math.Truncate(innumeric)).AddSeconds(Math.Truncate((innumeric % 1) * 24 * 60 * 60)).ToString("o");
                //  else
                // Feb 29th 1900 will show up as Mar 1st 1900 because Python won't handle that date
                //  return new DateTime(1899, 12, 31, 0, 0, 0) + timedelta(days=int(date), seconds=int((date % 1) * 24 * 60 * 60));
            }
        }
        public static string StringToCsv(string fromstr)
        {
            if (fromstr.Contains('"') || fromstr.Contains("\n") || fromstr.Contains(',') || fromstr.Contains("\t"))
                return "\"" + fromstr.Replace("\"", "\"\"") + "\"";
            return fromstr;
        }
    }
}

