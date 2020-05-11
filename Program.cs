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
      //string pathShStr = @"C:\user_main\python\FL_insurance_sample3 - Copy.xlsb\xl\sharedStrings.bin";
      //string pathWksht = @"C:\user_main\python\FL_insurance_sample3 - Copy.xlsb\xl\worksheets\sheet1.bin";
      string pathWksht = @"C:\user_main\python\xlsb\xl\worksheets\sheet1.bin";
      string pathStyle = @"C:\user_main\python\xlsb\xl\styles.bin";
      Dictionary<uint, string> shstr = new Dictionary<uint, string>();
      List<uint> datestyle = new List<uint>();
      try
      {
        using (FileStream fsSource = new FileStream(pathStyle, FileMode.Open, FileAccess.Read))
        {
          loadsharedstrings(fsSource, ref shstr);
        }
        //using (FileStream fsSource = new FileStream(pathShStr, FileMode.Open, FileAccess.Read))
        //{
        //  loadsharedstrings(fsSource, ref shstr);
        //}

        using (FileStream fsSource = new FileStream(pathWksht, FileMode.Open, FileAccess.Read))
        {
          readworksheet(fsSource, shstr);
        }
      }
      catch (FileNotFoundException ioEx)
      {
        Console.WriteLine(ioEx.Message);
      }
    }
    static public void loadstyles(FileStream fsSource, ref List<uint> DateStyles)
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
          case 44: // custom
            //[Black] [Green] [White] [Blue] [Magenta] [Yellow] [Cyan] [Red]
            break;
          case 47: // built-in
            break;
        }
      }
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
    static public void readworksheet(FileStream fsSource, Dictionary<uint, string> shstr)
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
              outputFile.Write("col {0} ", getcellno(data));
              outputFile.WriteLine("blank cell");
              break;
            case 2: // BrtCellRk
              outputFile.Write("col {0} ", getcellno(data));
              outputFile.WriteLine("rk number {0}", Convert.ToString(BitConverter.ToUInt32(data, 0), 2));
              //bit 0 - divide by 100 if 1
              //bit 1 - 30 sign. bits of float if 0 signed integer if 1
              break;
            case 3: // BrtCellError
              outputFile.Write("col {0} ", getcellno(data));
              outputFile.WriteLine("err {0}", data[8]);
              break;
            case 4: // BrtCellBool
              outputFile.Write("col {0} ", getcellno(data));
              outputFile.WriteLine("bool {0}", data[8]);
              break;
            case 5: // BrtCellReal
              outputFile.Write("col {0} ", getcellno(data));
              outputFile.WriteLine("dbl {0}", BitConverter.ToDouble(data, 8));
              break;
            case 6: // BrtCellSt
              outputFile.Write("col {0} ", getcellno(data));
              outputFile.WriteLine("value {0}", getxlwidestring(data, 8));
              break;
            case 7: // BrtCellIsst
              outputFile.Write("col {0} ", getcellno(data));
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
    static public uint getcellno(byte[] buffer)
    {
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
  }
}

