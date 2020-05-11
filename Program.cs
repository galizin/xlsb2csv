using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace xlsbtocsv
{
  class Program
  {
    static void Main(string[] args)
    {
      //2.2.1
      //2.1.8
      string pathShStr = @"C:\user_main\python\FL_insurance_sample3 - Copy.xlsb\xl\sharedStrings.bin";
      string pathWksht = @"C:\user_main\python\FL_insurance_sample3 - Copy.xlsb\xl\worksheets\sheet1.bin";
      string pathStyle = @"C:\user_main\python\FL_insurance_sample3 - Copy.xlsb\xl\styles.bin";
      //string pathWksht = @"C:\user_main\python\xlsb\xl\worksheets\sheet1.bin";
      //string pathStyle = @"C:\user_main\python\xlsb\xl\styles.bin";
      Dictionary<uint, string> shstr = new Dictionary<uint, string>();
      List<uint> datestyle = new List<uint>();
      try
      {
        using (FileStream fsSource = new FileStream(pathStyle, FileMode.Open, FileAccess.Read))
        {
          loadstyles(fsSource, ref datestyle);
        }
        using (FileStream fsSource = new FileStream(pathShStr, FileMode.Open, FileAccess.Read))
        {
          loadsharedstrings(fsSource, ref shstr);
        }
        using (FileStream fsSource = new FileStream(pathWksht, FileMode.Open, FileAccess.Read))
        {
          readworksheet(fsSource, shstr, datestyle);
        }
      }
      catch (FileNotFoundException ioEx)
      {
        Console.WriteLine(ioEx.Message);
      }
    }
    static public void loadstyles(FileStream fsSource, ref List<uint> outDateStyles)
    {
      //using (StreamWriter outputFile = new StreamWriter("output_styl.txt", false, Encoding.UTF8))
      //{
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
            //[Black] [Green] [White] [Blue] [Magenta] [Yellow] [Cyan] [Red]
            if (filteredstring.Contains("y") || filteredstring.Contains("d") || filteredstring.Contains("h") || filteredstring.Contains("m") || filteredstring.Contains("s") || filteredstring.Contains("a") || filteredstring.Contains("p"))
            {
              //outputFile.Write("date ");
              datestyles.Add(BitConverter.ToUInt16(data, 0));
            }
            //outputFile.WriteLine("custom style {0} {1}", BitConverter.ToUInt16(data, 0), getxlwidestring(data, 2));
            break;
          case 47: // BrtXF
            //14-22,45-47
            //outputFile.WriteLine("builtin style parent {0} id {1}", BitConverter.ToUInt16(data, 0), BitConverter.ToUInt16(data, 2));
            xf.Add(styleid, BitConverter.ToUInt16(data, 2));
            styleid++;
            break;
          //case 48: // style info              
          //  outputFile.WriteLine("style info id {0} builtin bit {1} no {2} level {3}", BitConverter.ToUInt16(data, 0), data[4]%1u, data[6], data[7]);
          //  break;
        }
      }
      for (uint i = 0; i < styleid; i++)
      {
        if (datestyles.Contains(xf[i]))
          outDateStyles.Add(i);
      }
      //}
    }
    static public void loadsharedstrings(FileStream fsSource, ref Dictionary<uint, string> shstr)
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
    static public void readworksheet(FileStream fsSource, Dictionary<uint, string> shstr, List<uint> datestyles)
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
              //outputFile.WriteLine("rk number {0} {1}{2}{3}{4} {5}", x, Convert.ToString(data[8], 16), Convert.ToString(data[9], 16), Convert.ToString(data[10], 16), Convert.ToString(data[11], 16), Convert.ToString(BitConverter.ToUInt32(data, 8), 2));
              if (dateformatted(data, datestyles))
                outputFile.WriteLine("rk date {0}", stringdate(x));
              else
                outputFile.WriteLine("rk {0}", x);
              //bit 0 - divide by 100 if 1
              //bit 1 - 30 sign. bits of float if 0 signed integer if 1
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
          //BitConverter:
          //ToUInt32(Byte[], Int32)
          //string s = System.Text.Encoding.UTF8.GetString(buffer, 0, buffer.Length);

          //RichStr: skip 1 byte then XlWideString

          //XLWideString: 4 bytes length, then unicode chars

          //DATACELL = CELLMETA (BrtCellBlank / BrtCellRk / BrtCellError / BrtCellBool / BrtCellReal / BrtCellIsst / BrtCellSt) 
          //FMLACELL = CELLMETA (BrtFmlaString / BrtFmlaNum / BrtFmlaBool / BrtFmlaError)
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
      //if (datastyles.Contains(styleid))
      //f.Write("date ");
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
    static public void readrecord(out int rec_id, ref byte[] data, FileStream fsSource)
    {
      rec_id = read_id(fsSource);
      if (rec_id == -1)
        return;
      int rec_len = read_len(fsSource);
      data = new byte[rec_len];
      fsSource.Read(data, 0, rec_len);
    }
    static public int read_id(FileStream fsSource)
    {
      //p186 record info
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
    static public int read_len(FileStream fsSource)
    {
      //2147483647 - upper bound
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

