using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Threading;
 
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using DataMatrix.net;
namespace Server_print
{
    class Program
    {
        public static Excel.Application excelapp;
        public static object _lock = new object();

        public static string[,] arrData = new string[500000, 6];
        public static int LastRow = 0;
        DataMatrix.net.DmtxImageEncoder encoder = new DataMatrix.net.DmtxImageEncoder();

        static void Main(string[] args)
        {
            // для генерации datamatrix кодов

            Console.OutputEncoding = Encoding.GetEncoding(866);
            excelapp = new Excel.Application();
            excelapp.Visible = false;
            excelapp.Workbooks.Open(@AppDomain.CurrentDomain.BaseDirectory + "\\Ready_base.xlsx");
            Excel.Workbook WB = excelapp.ActiveWorkbook;
            Excel.Worksheet WS = excelapp.ActiveSheet;
            int LastR = WB.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            LastRow = LastR;
            //MessageBox.Show(Convert.ToString(LastRow));
            // A это штрихкод, B это артикул

            // Сначала ищем по артикулу для первого шаблона

            //string Obj = String.Empty;
            //string Name = String.Empty;

            var arrData2 = (object[,])WS.Range["A1:E" + LastR].Value;
            WB.Close(false);
            excelapp.Quit();
            for (int i = 1; i <= LastRow; i++)
            {
                arrData[i, 1] = Convert.ToString(arrData2[i, 1]);
                arrData[i, 2] = Convert.ToString(arrData2[i, 2]);
                arrData[i, 3] = Convert.ToString(arrData2[i, 3]);
                arrData[i, 4] = Convert.ToString(arrData2[i, 4]);
                arrData[i, 5] = Convert.ToString(arrData2[i, 5]);
            }

            //string[,] arrLast_b = new string[20000, 4];



            TcpListener server = null;
            //DataMatrix.net.DmtxImageEncoder encoder = new DataMatrix.net.DmtxImageEncoder();
            try
            {
                // Узнаю свой IP
                string IP = Dns.GetHostEntry(Dns.GetHostName()).AddressList.First(address => address.AddressFamily == AddressFamily.InterNetwork).ToString();

                // Устанавливаю количество потоков мин и макс
                int MaxThreadsCount = Environment.ProcessorCount * 4;
                ThreadPool.SetMaxThreads(MaxThreadsCount, MaxThreadsCount);
                ThreadPool.SetMinThreads(2, 2);

                // Устанавливаю порт и IP сервера
                Int32 port = 9595;
                IPAddress localAddr = IPAddress.Parse(IP);

                //Счетчик соединений
                int counter = 0;

                server = new TcpListener(localAddr, port);

                Console.WriteLine("Configure MultithreadServer:");
                Console.WriteLine(" IP Address: " + IP.ToString());
                Console.WriteLine(" Port: " + port.ToString());
                Console.WriteLine("Threads: " + MaxThreadsCount.ToString());
                Console.WriteLine("\nServer is run!\n");
                Console.WriteLine("Configure MultithreadServer:");

                server.Start();
                while (true)
                {
                    //Отдельный поток каждому клиенту
                    Console.WriteLine("\nWaiting for connection...");
                    ThreadPool.QueueUserWorkItem(ClientProccesing, server.AcceptTcpClient());
                    counter++;
                    Console.WriteLine("\nConnection №" + counter.ToString() + "!");
                }

            }
            catch (SocketException e)
            {
                Console.WriteLine("SocketException: {0}", e);
            }
            finally
            {
                server.Stop();
            }
            Console.WriteLine("\nPress Enter...");
            Console.Read();

        }


        static void ClientProccesing(object client_obj)
        {
            DataMatrix.net.DmtxImageEncoder encoder = new DataMatrix.net.DmtxImageEncoder();

            Byte[] bytes = new Byte[256];
            String data = null;
            TcpClient client = client_obj as TcpClient;
            data = null;
            NetworkStream stream = client.GetStream();
            int i, j, ArrayCount, c = 0, cs = 0;
            string[] STR_all_elements = new string[101];
            string Cennik = "";
            for (i = 0; i < 101; i++)
            {
                STR_all_elements[i] = "";
            }
            i = 0;
            while ((j = stream.Read(bytes, 0, bytes.Length)) != 0)
            {

                data = System.Text.Encoding.ASCII.GetString(bytes, 0, j);
                STR_all_elements[i] = data;
                //data = "Привет!";
                if (Convert.ToInt32(STR_all_elements[i].Substring(STR_all_elements[i].Length - 1, 1)) == 9)
                {
                    if (Convert.ToInt32(STR_all_elements[i].Substring(STR_all_elements[i].Length - 2, 1)) == 1)
                    {
                        //MessageBox.Show("по артикулу", "Внимание");
                        Cennik = STR_all_elements[i].Substring(0, STR_all_elements[i].Length - 5);
                        for (j = 1; j < LastRow; j++)
                        {
                            if (Cennik == Convert.ToString(arrData[j, 2]))
                            {
                                data = Convert.ToString(arrData[j, 2]) + " " + Convert.ToString(arrData[j, 3]) + " " + Convert.ToString(arrData[j, 4]) + "руб.";
                                c = 1;
                                break;
                            }
                        }
                        cs = 1;
                    }
                    else
                    {
                        //MessageBox.Show("по штрихкоду", "Внимание");
                        Cennik = STR_all_elements[i].Substring(0, STR_all_elements[i].Length - 5);
                        for (j = 1; j < LastRow; j++)
                        {
                            if (Cennik == Convert.ToString(arrData[j, 1]))
                            {
                                //получаю артикул
                                data = Convert.ToString(arrData[j, 2]) + " " + Convert.ToString(arrData[j, 3]) + " " + Convert.ToString(arrData[j, 4]) + "руб.";
                                c = 1;
                                break;

                            }
                        }
                        cs = 1;

                    }

                }
                if (cs == 1 && c == 0)
                {
                    data = "Товар в базе не найден!";
                }
                Console.WriteLine(STR_all_elements[i].ToString() + "\n");
                byte[] msg = System.Text.Encoding.UTF8.GetBytes(data);
                stream.Write(msg, 0, msg.Length);
                i++;
            }
            ArrayCount = i;
            client.Close();
            if (c == 1 || cs == 1)
            {
                goto m;
            }

            for (i = 0; i < ArrayCount; i++)
            {
                //предотвращение выбивания сервера при получении пустого элемента
                if (STR_all_elements[i] == null)
                    STR_all_elements[i] = " ";
            }


            string[] STR_art_1 = new string[100];
            string[] STR_art_2 = new string[100];
            string[] STR_art_3 = new string[100];
            string[] STR_shtrih_1 = new string[100];
            string[] STR_shtrih_2 = new string[100];
            string[] STR_shtrih_3 = new string[100];

            // Массивы с данными для печати (первый номер - ноиер товара, второй его Артикул, наименование, цена)
            string[,] Shablon_1 = new string[100, 4];
            string[,] Shablon_2 = new string[100, 4];
            string[,] Shablon_3 = new string[100, 4];

            int k1 = 0, k2 = 0, k3 = 0, j1 = 0, j2 = 0, j3 = 0, j4 = 0, j5 = 0, j6 = 0;
            //заполняем массив пустыми ячейками
            for (i = 0; i < 100; i++)
            {
                STR_art_1[i] = "";
                STR_art_2[i] = "";
                STR_art_3[i] = "";
                STR_shtrih_1[i] = "";
                STR_shtrih_2[i] = "";
                STR_shtrih_3[i] = "";
                for (j = 0; j < 4; j++)
                {
                    Shablon_1[i, j] = "";
                    Shablon_2[i, j] = "";
                    Shablon_3[i, j] = "";
                }
            }


            for (i = 0; i < ArrayCount; i++)
            {

                //работаем с массивом
                if (Convert.ToInt32(STR_all_elements[i].Substring(STR_all_elements[i].Length - 1, 1)) == 1)
                {
                    //шаблон 1
                    if (Convert.ToInt32(STR_all_elements[i].Substring(STR_all_elements[i].Length - 2, 1)) == 1)
                    {
                        //MessageBox.Show("по артикулу", "Внимание");
                        STR_art_1[j1] = STR_all_elements[i].Substring(0, STR_all_elements[i].Length - 5);
                        j1++;
                    }
                    else
                    {
                        //MessageBox.Show("по штрихкоду", "Внимание");
                        STR_shtrih_1[j2] = STR_all_elements[i].Substring(0, STR_all_elements[i].Length - 5);
                        j2++;
                    }
                }
                else if (Convert.ToInt32(STR_all_elements[i].Substring(STR_all_elements[i].Length - 1, 1)) == 2)
                {
                    //шаблон 2
                    if (Convert.ToInt32(STR_all_elements[i].Substring(STR_all_elements[i].Length - 2, 1)) == 1)
                    {
                        //MessageBox.Show("по артикулу", "Внимание");
                        STR_art_2[j3] = STR_all_elements[i].Substring(0, STR_all_elements[i].Length - 5);
                        j3++;
                    }
                    else
                    {
                        //MessageBox.Show("по штрихкоду", "Внимание");
                        STR_shtrih_2[j4] = STR_all_elements[i].Substring(0, STR_all_elements[i].Length - 5);
                        j4++;
                    }
                }
                else if (Convert.ToInt32(STR_all_elements[i].Substring(STR_all_elements[i].Length - 1, 1)) == 3)
                {
                    //шаблон 3
                    if (Convert.ToInt32(STR_all_elements[i].Substring(STR_all_elements[i].Length - 2, 1)) == 1)
                    {
                        //MessageBox.Show("по артикулу", "Внимание");
                        STR_art_3[j5] = STR_all_elements[i].Substring(0, STR_all_elements[i].Length - 5);
                        j5++;
                    }
                    else
                    {
                        //MessageBox.Show("по штрихкоду", "Внимание");
                        STR_shtrih_3[j6] = STR_all_elements[i].Substring(0, STR_all_elements[i].Length - 5);
                        j6++;
                    }
                }

            }
            // данные отсортированы по типам в массивы
            string FindObj = String.Empty;
            for (i = 0; i < j1; i++)
            {
                FindObj = STR_art_1[i];
                for (j = 1; j < LastRow; j++)
                {
                    if (FindObj == Convert.ToString(arrData[j, 2]))
                    {
                        //получаю артикул

                        Shablon_1[k1, 0] = Convert.ToString(arrData[j, 2]);

                        Bitmap bmp = encoder.EncodeImage(Shablon_1[k1, 0].ToString());
                        bmp.Save(@AppDomain.CurrentDomain.BaseDirectory + "\\DataM\\" + Shablon_1[k1, 0].ToString() + ".Png", System.Drawing.Imaging.ImageFormat.Png);
                        //получаю наименование
                        Shablon_1[k1, 1] = Convert.ToString(arrData[j, 3]);
                        //получаю цену
                        Shablon_1[k1, 2] = Convert.ToString(arrData[j, 4]);
                        //контроль Долго хранимого товара
                        Shablon_1[k1, 3] = Convert.ToString(arrData[j, 5]);
                        k1++;
                        break;
                    }
                }


            }
            // Теперь по штрихкоду для первого шаблона
            for (i = 0; i < j2; i++)
            {
                FindObj = STR_shtrih_1[i];
                for (j = 1; j < LastRow; j++)
                {
                    if (FindObj == Convert.ToString(arrData[j, 1]))
                    {
                        //получаю артикул
                        Shablon_1[k1, 0] = Convert.ToString(arrData[j, 2]);
                        try
                        {
                            Bitmap bmp = encoder.EncodeImage(Shablon_1[k1, 0].ToString());
                            bmp.Save(@AppDomain.CurrentDomain.BaseDirectory + "\\DataM\\" + Shablon_1[k1, 0].ToString() + ".Png", System.Drawing.Imaging.ImageFormat.Png);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Exception:", e);
                        }


                        //получаю наименование
                        Shablon_1[k1, 1] = Convert.ToString(arrData[j, 3]);
                        //получаю цену
                        Shablon_1[k1, 2] = Convert.ToString(arrData[j, 4]);
                        //контроль Долго хранимого товара
                        Shablon_1[k1, 3] = Convert.ToString(arrData[j, 5]);
                        k1++;
                        break;
                    }
                }
            }
            //Шаблон 2

            for (i = 0; i < j3; i++)
            {
                FindObj = STR_art_2[i];
                //MessageBox.Show(STR_art_1[i]);
                for (j = 1; j < LastRow; j++)
                {
                    if (FindObj == Convert.ToString(arrData[j, 2]))
                    {
                        //получаю артикул
                        Shablon_2[k2, 0] = Convert.ToString(arrData[j, 2]);
                        Bitmap bmp = encoder.EncodeImage(Shablon_2[k2, 0].ToString());
                        bmp.Save(@AppDomain.CurrentDomain.BaseDirectory + "\\DataM\\" + Shablon_2[k2, 0].ToString() + ".Png", System.Drawing.Imaging.ImageFormat.Png);

                        //получаю наименование
                        Shablon_2[k2, 1] = Convert.ToString(arrData[j, 3]);
                        //получаю цену
                        Shablon_2[k2, 2] = Convert.ToString(arrData[j, 4]);
                        //контроль Долго хранимого товара
                        Shablon_2[k2, 3] = Convert.ToString(arrData[j, 5]);
                        k2++;
                        break;
                    }
                }


            }
            // Теперь по штрихкоду для второго шаблона

            for (i = 0; i < j4; i++)
            {
                FindObj = STR_shtrih_2[i];
                for (j = 1; j < LastRow; j++)
                {
                    if (FindObj == Convert.ToString(arrData[j, 1]))
                    {
                        //получаю артикул
                        Shablon_2[k2, 0] = Convert.ToString(arrData[j, 2]);
                        Bitmap bmp = encoder.EncodeImage(Shablon_2[k2, 0].ToString());
                        bmp.Save(@AppDomain.CurrentDomain.BaseDirectory + "\\DataM\\" + Shablon_2[k2, 0].ToString() + ".Png", System.Drawing.Imaging.ImageFormat.Png);

                        //получаю наименование
                        Shablon_2[k2, 1] = Convert.ToString(arrData[j, 3]);
                        //получаю цену
                        Shablon_2[k2, 2] = Convert.ToString(arrData[j, 4]);
                        //контроль Долго хранимого товара
                        Shablon_2[k2, 3] = Convert.ToString(arrData[j, 5]);
                        k2++;
                        break;
                    }
                }


            }
            //Шаблон 3

            for (i = 0; i < j5; i++)
            {
                FindObj = STR_art_3[i];
                for (j = 1; j < LastRow; j++)
                {

                    if (FindObj == Convert.ToString(arrData[j, 2]))
                    {
                        //получаю артикул
                        Shablon_3[k3, 0] = Convert.ToString(arrData[j, 2]);
                        Bitmap bmp = encoder.EncodeImage(Shablon_3[k3, 0].ToString());
                        bmp.Save(@AppDomain.CurrentDomain.BaseDirectory + "\\DataM\\" + Shablon_3[k3, 0].ToString() + ".Png", System.Drawing.Imaging.ImageFormat.Png);

                        //получаю наименование
                        Shablon_3[k3, 1] = Convert.ToString(arrData[j, 3]);
                        //получаю цену
                        Shablon_3[k3, 2] = Convert.ToString(arrData[j, 4]);
                        //контроль Долго хранимого товара
                        Shablon_3[k3, 3] = Convert.ToString(arrData[j, 5]);
                        k3++;
                        break;
                    }
                }


            }
            // Теперь по штрихкоду для первого шаблона

            for (i = 0; i < j6; i++)
            {
                FindObj = STR_shtrih_3[i];
                //MessageBox.Show(STR_art_1[i]);
                for (j = 1; j < LastRow; j++)
                {
                    if (FindObj == Convert.ToString(arrData[j, 1]))
                    {
                        //получаю артикул
                        Shablon_3[k3, 0] = Convert.ToString(arrData[j, 2]);
                        Bitmap bmp = encoder.EncodeImage(Shablon_3[k3, 0].ToString());
                        bmp.Save(@AppDomain.CurrentDomain.BaseDirectory + "\\DataM\\" + Shablon_3[k3, 0].ToString() + ".Png", System.Drawing.Imaging.ImageFormat.Png);

                        //получаю наименование
                        Shablon_3[k3, 1] = Convert.ToString(arrData[j, 3]);
                        //получаю цену
                        Shablon_3[k3, 2] = Convert.ToString(arrData[j, 4]);
                        //контроль Долго хранимого товара
                        Shablon_3[k3, 3] = Convert.ToString(arrData[j, 5]);
                        k3++;
                        break;
                    }
                }
            }

            //массивы готовы к печати. Shablon_1 Shablon_2 Shablon_3
            // Массивы с данными для печати (первый номер - товар, второй его Артикул, наименование, цена) счетчики k1, k2, k3


            lock (_lock)
            {
                //Console.WriteLine("Thread " + Thread.CurrentThread.ManagedThreadId);
                int x, y, z, t, x1, y1, z1, t1;
                Excel.Workbook WB = excelapp.ActiveWorkbook;
                Excel.Worksheet WS = excelapp.ActiveSheet;
                if (k1 > 0)
                {
                    excelapp.Workbooks.Open(@AppDomain.CurrentDomain.BaseDirectory + "\\Sh_1.xlsx");
                    WB = excelapp.ActiveWorkbook;
                    WS = excelapp.ActiveSheet;
                    excelapp.Visible = false;
                    Excel.Range rg = null;

                    float il, it, iw, ih;
                    float zExcelPixel = 0.746835443f;

                    x = 1;
                    y = 3;
                    z = 9;
                    t = 1;
                    x1 = 11;
                    y1 = 1;
                    z1 = 3;
                    t1 = 8;
                    //Console.WriteLine("Thread st 2" + Thread.CurrentThread.ManagedThreadId);
                    for (i = 0; i < k1; i++)
                    {
                        WS.Cells[x, x1] = Shablon_1[i, 0];
                        WS.Cells[y, y1] = Shablon_1[i, 1];
                        WS.Cells[z, z1] = Shablon_1[i, 2];
                        Image im = Image.FromFile(@AppDomain.CurrentDomain.BaseDirectory + "\\DataM\\" + Shablon_1[i, 0].ToString() + ".Png");


                        // размеры изображения для Shape нужно преобразовывать
                        iw = zExcelPixel * im.Width;// получаем из ширины исходного изображения
                        ih = zExcelPixel * im.Height;
                        if (i % 2 == 0)
                        {
                            rg = WS.get_Range("M" + z, "M" + z);
                        }
                        else
                        {
                            rg = WS.get_Range("AA" + z, "AA" + z);
                        }
                        il = (float)(double)rg.Left;// размеры поступают в double упакованый в object
                        it = (float)(double)rg.Top;
                        WS.Shapes.AddPicture(@AppDomain.CurrentDomain.BaseDirectory + "\\DataM\\" + Shablon_1[i, 0].ToString() + ".Png", Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, il, it, iw - 20, ih - 20);
                        im.Dispose();
                        if (Shablon_1[i, 3] == "1")
                        {
                            WS.Cells[t, t1] = ".";
                        }

                        if (i % 2 == 0)
                        {
                            x1 = 25;
                            y1 = 15;
                            z1 = 17;
                            t1 = 22;
                        }
                        else
                        {
                            x1 = 11;
                            y1 = 1;
                            z1 = 3;
                            t1 = 8;

                            x = x + 11;
                            y = y + 11;
                            z = z + 11;
                            t = t + 11;
                        }

                        //Console.WriteLine("Thread st 3 " + Thread.CurrentThread.ManagedThreadId);
                    }
                    if (k1 <= 10)
                    {
                        WS.PrintOutEx(1, 1);
                    }
                    else if (k1 > 10 && k1 <= 20)
                    {
                        WS.PrintOutEx(1, 2);
                    }
                    else if (k1 > 20 && k1 <= 30)
                    {
                        WS.PrintOutEx(1, 3);
                    }
                    else if (k1 > 30 && k1 <= 40)
                    {
                        WS.PrintOutEx(1, 4);
                    }
                    else if (k1 > 40 && k1 <= 50)
                    {
                        WS.PrintOutEx(1, 5);
                    }
                    else if (k1 > 50 && k1 <= 60)
                    {
                        WS.PrintOutEx(1, 6);
                    }
                    else if (k1 > 60 && k1 <= 70)
                    {
                        WS.PrintOutEx(1, 7);
                    }
                    else if (k1 > 70 && k1 <= 80)
                    {
                        WS.PrintOutEx(1, 8);
                    }
                    else if (k1 > 80 && k1 <= 90)
                    {
                        WS.PrintOutEx(1, 9);
                    }
                    else if (k1 > 90 && k1 <= 100)
                    {
                        WS.PrintOutEx(1, 10);
                    }

                    GC.Collect();
                    WB.Close(false);
                    excelapp.Quit();


                }
                //Console.WriteLine("Thread st4" + Thread.CurrentThread.ManagedThreadId);
                //печать 2го шаблона

                if (k2 > 0)
                {
                    excelapp.Workbooks.Open(@AppDomain.CurrentDomain.BaseDirectory + "\\Sh_2.xlsx");
                    WB = excelapp.ActiveWorkbook;
                    WS = excelapp.ActiveSheet;
                    excelapp.Visible = false;
                    Excel.Range rg = null;

                    float il, it, iw, ih;
                    float zExcelPixel = 0.746835443f;

                    x = 1;
                    y = 2;
                    z = 6;
                    t = 1;

                    x1 = 8;
                    y1 = 1;
                    z1 = 2;
                    t1 = 5;

                    for (i = 0; i < k2; i++)
                    {
                        WS.Cells[x, x1] = Shablon_2[i, 0];
                        WS.Cells[y, y1] = Shablon_2[i, 1];
                        WS.Cells[z, z1] = Shablon_2[i, 2];
                        Image im = Image.FromFile(@AppDomain.CurrentDomain.BaseDirectory + "\\DataM\\" + Shablon_2[i, 0].ToString() + ".Png");


                        // размеры изображения для Shape нужно преобразовывать
                        iw = zExcelPixel * im.Width;// получаем из ширины исходного изображения
                        ih = zExcelPixel * im.Height;
                        if (i % 2 == 0)
                        {
                            rg = WS.get_Range("I" + z, "I" + z);
                        }
                        else
                        {
                            rg = WS.get_Range("S" + z, "S" + z);
                        }
                        il = (float)(double)rg.Left;// размеры поступают в double упакованый в object
                        it = (float)(double)rg.Top;
                        WS.Shapes.AddPicture(@AppDomain.CurrentDomain.BaseDirectory + "\\DataM\\" + Shablon_2[i, 0].ToString() + ".Png", Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, il, it, iw - 20, ih - 20);
                        im.Dispose();




                        if (Shablon_2[i, 3] == "1")
                        {
                            WS.Cells[t, t1] = ".";
                        }

                        //правая левая стороны листа
                        if (i % 2 == 0)
                        {
                            //координаты ячеек
                            x1 = 18;
                            y1 = 11;
                            z1 = 12;
                            t1 = 15;
                        }
                        else
                        {
                            x1 = 8;
                            y1 = 1;
                            z1 = 2;
                            t1 = 5;

                            x = x + 8;
                            y = y + 8;
                            z = z + 8;
                            t = t + 8;
                        }

                    }
                    if (k2 <= 10)
                    {
                        WS.PrintOutEx(1, 1);
                    }
                    else if (k2 > 10 && k2 <= 20)
                    {
                        WS.PrintOutEx(1, 2);
                    }
                    else if (k2 > 20 && k2 <= 30)
                    {
                        WS.PrintOutEx(1, 3);
                    }
                    else if (k2 > 30 && k2 <= 40)
                    {
                        WS.PrintOutEx(1, 4);
                    }
                    else if (k2 > 40 && k2 <= 50)
                    {
                        WS.PrintOutEx(1, 5);
                    }
                    else if (k2 > 50 && k2 <= 60)
                    {
                        WS.PrintOutEx(1, 6);
                    }
                    else if (k2 > 60 && k2 <= 70)
                    {
                        WS.PrintOutEx(1, 7);
                    }
                    else if (k2 > 70 && k2 <= 80)
                    {
                        WS.PrintOutEx(1, 8);
                    }
                    else if (k2 > 80 && k2 <= 90)
                    {
                        WS.PrintOutEx(1, 9);
                    }
                    else if (k2 > 90 && k2 <= 100)
                    {
                        WS.PrintOutEx(1, 10);
                    }
                    WB.Close(false);
                    excelapp.Quit();

                }
                //3й шаблон
                if (k3 > 0)
                {
                    excelapp.Workbooks.Open(@AppDomain.CurrentDomain.BaseDirectory + "\\Sh_3.xlsx");
                    WB = excelapp.ActiveWorkbook;
                    WS = excelapp.ActiveSheet;
                    excelapp.Visible = false;
                    Excel.Range rg = null;

                    float il, it, iw, ih;
                    float zExcelPixel = 0.746835443f;
                    x = 1;
                    y = 2;
                    z = 8;
                    t = 1;

                    x1 = 10;
                    y1 = 1;
                    z1 = 2;
                    t1 = 7;
                    for (i = 0; i < k3; i++)
                    {
                        WS.Cells[x, x1] = Shablon_3[i, 0];
                        WS.Cells[y, y1] = Shablon_3[i, 1];
                        WS.Cells[z, z1] = Shablon_3[i, 2];
                        Image im = Image.FromFile(@AppDomain.CurrentDomain.BaseDirectory + "\\DataM\\" + Shablon_3[i, 0].ToString() + ".Png");


                        // размеры изображения для Shape нужно преобразовывать
                        iw = zExcelPixel * im.Width;// получаем из ширины исходного изображения
                        ih = zExcelPixel * im.Height;
                        if (i % 2 == 0)
                        {
                            rg = WS.get_Range("K" + (z + 1), "K" + (z + 1));
                        }
                        else
                        {
                            rg = WS.get_Range("W" + (z + 1), "W" + (z + 1));
                        }
                        il = (float)(double)rg.Left;// размеры поступают в double упакованый в object
                        it = (float)(double)rg.Top;
                        WS.Shapes.AddPicture(@AppDomain.CurrentDomain.BaseDirectory + "\\DataM\\" + Shablon_3[i, 0].ToString() + ".Png", Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, il, it, iw - 20, ih - 20);
                        im.Dispose();

                        if (Shablon_3[i, 3] == "1")
                        {
                            WS.Cells[t, t1] = ".";
                        }
                        //правая левая стороны листа
                        if (i % 2 == 0)
                        {
                            //координаты ячеек
                            x1 = 22;
                            y1 = 13;
                            z1 = 14;
                            t1 = 19;
                        }
                        else
                        {
                            x1 = 10;
                            y1 = 1;
                            z1 = 2;
                            t1 = 7;

                            x = x + 11;
                            y = y + 11;
                            z = z + 11;
                            t = t + 11;
                        }
                    }

                    if (k3 <= 10)
                    {
                        WS.PrintOutEx(1, 1);
                    }
                    else if (k3 > 10 && k3 <= 20)
                    {
                        WS.PrintOutEx(1, 2);
                    }
                    else if (k3 > 20 && k3 <= 30)
                    {
                        WS.PrintOutEx(1, 3);
                    }
                    else if (k3 > 30 && k3 <= 40)
                    {
                        WS.PrintOutEx(1, 4);
                    }
                    else if (k3 > 40 && k3 <= 50)
                    {
                        WS.PrintOutEx(1, 5);
                    }
                    else if (k3 > 50 && k3 <= 60)
                    {
                        WS.PrintOutEx(1, 6);
                    }
                    else if (k3 > 60 && k3 <= 70)
                    {
                        WS.PrintOutEx(1, 7);
                    }
                    else if (k3 > 70 && k3 <= 80)
                    {
                        WS.PrintOutEx(1, 8);
                    }
                    else if (k3 > 80 && k3 <= 90)
                    {
                        WS.PrintOutEx(1, 9);
                    }
                    else if (k3 > 90 && k3 <= 100)
                    {
                        WS.PrintOutEx(1, 10);
                    }
                    WB.Close(false);
                    excelapp.Quit();
                    DirectoryInfo dirInfo = new DirectoryInfo(@AppDomain.CurrentDomain.BaseDirectory + "\\DataM\\");
                    foreach (FileInfo file in dirInfo.GetFiles())
                    {
                        file.Delete();
                    }
                    //Console.WriteLine("Exit" + Thread.CurrentThread.ManagedThreadId);
                }
            }
            m:;
        }
    }
}

