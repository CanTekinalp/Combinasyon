using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;

namespace combinasyon
{
    class Program
    {

        public static List<List<string>> data = new List<List<string>>();

        //Satır boyunca en az bir adet data mevcut mu kontrol ediyoruz.
        public static bool AnyData(IExcelDataReader reader, int groupSize)
        {
            for (int i = 0; i < groupSize; i++)
            {
                if (reader.GetString(i) != null)
                    return true;
            }
            return false;
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Dosyanın yolunu , adı ve uzantısı ile birlikte girin :");
            Console.WriteLine("Örneğin C:\\Users\\Can\\Desktop\\excel.xlsx (Sadece ilk sayfayı okuyacak)");

            string dir = Console.ReadLine();
            

            FileStream fs = File.Open(dir, FileMode.Open);
            IExcelDataReader reader = ExcelReaderFactory.CreateReader(fs);

            //Tekil Grup Kodları hücresini okuyoruz.
            reader.Read();
            //Bir alt satıra geçiyoruz.
            reader.Read();

            int groupSize = 0;

            bool endOfGroup = false;

            //Satır boyunca ilerliyoruz.
            while (!endOfGroup)
            {
                if (reader.GetString(groupSize) != null)
                {
                    groupSize++;
                }
                else
                {
                    //Sonraki col da boş ise datanın sonuna gelmiş oluyoruz.
                    if (reader.GetString(groupSize + 1) == null)
                        endOfGroup = true;
                    else
                        groupSize++;

                }

            }

            //Grup sayısı kadar liste oluşturuyoruz.
            for (int i = 0; i < groupSize; i++)
            {
                data.Add(new List<string>());
            }

            //Bir alt satıra geçiyoruz , grup verilerini okumaya başlıyoruz.
            reader.Read();

            //Satır boyunca verileri okuyoruz , eğer data mevcutsa listeye ekliyoruz.
            while (AnyData(reader, groupSize))
            {
                for (int i = 0; i < groupSize; i++)
                {
                    if (reader.GetString(i) != null)
                        data[i].Add(reader.GetString(i));

                }
                reader.Read();

            }

            //Eleman sayısı sıfır olan özelliklere bir tane boş string atıyoruz.
            for (int i = 0; i < data.Count; i++)
            {
                if (data[i].Count == 0)
                    data[i].Add(string.Empty);
            }

            long[] div = new long[data.Count];
            long totalDiv = 1;

            for (int i = data.Count - 1; i >= 0; i--)
            {
                div[i] = totalDiv;
                totalDiv *= data[i].Count;
            }

            for (long combo = 0; combo < totalDiv; combo++)
            {
                for (int i = 0; i < data.Count; i++)
                {
                    int digit = (int)(combo / div[i] % data[i].Count);

                    if (i != 0)
                        Console.Write(' ');
                    Console.Write(data[i][digit]);
                }
                Console.WriteLine();
            }

            //Dosyayı kapatıyoruz.
            fs.Close();

            Console.Read();
        }
    }
}
