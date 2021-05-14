using System;
using GemBox.Spreadsheet;

namespace ConsoleApp
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Product product = new Product();
            
            String[] questions = new string[]
            {
                "наименование изделия",
                "код изделия",  
                "единицу измерения",
                "цену за единицу измеряемого изделия",
                "номер цеха",
                "номер склада",
                "наименование склада",
                "номер накладной",
                "дату договора",
                "дату в накладной",
                "количество товара",
                "номер договора"
            };
            
            
            for (int iterator = 0; iterator < questions.Length; iterator++)
            {
                //iterator = 8
                Console.Write("Введите " + questions[3] + ": ");
                bool status = product.SetParameter(iterator, Console.ReadLine());
                if (status == false)
                {
                    iterator--;
                }
            }
            
            String[] personQuestions = new string[]
            {
                "имя",
                "фамилию",
                "город",
                "номер телефона"
            };
            
            for (int i = 0; i < personQuestions.Length; i++)
            {
                Console.Write("Введите " + personQuestions[i] + " поставщика: ");
                bool status = product.Diler.SetParameter(i, Console.ReadLine());
                if (status == false) i--;
            }
            
            for (int i = 0; i < personQuestions.Length; i++)
            {
                Console.Write("Введите " + personQuestions[i] + " покупателя: ");
                bool status = product.Buyer.SetParameter(i, Console.ReadLine());
                if (status == false) i--;
            }


            // product.Name = "Продукт";
            // product.Code = "ABC";
            // product.MeasurementUnit = "см";
            // product.UnitPrice = 3000;
            // product.FabricId = 15;
            // product.StorageNumber = 18;
            // product.StorageName = "Склат";
            // product.Invoice = 939;
            // product.ContractDate = new DateTime(2002, 12, 31);
            // product.InvoiceDate = new DateTime(2001, 12, 30);
            // product.Amount = 10;
            // product.ContractNumber = 500;
            // product.Diler = new Person("ПостИМЯ","ПОСТФАМ","САМАРА","ТЕЛЕФОН");
            // product.Buyer = new Person("ПокИМЯ","ПокФАМ","ПИТЕР","ТЕЛЕФОН2");
            
            
            
            
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile excelFile = new ExcelFile();

            ExcelWorksheet mainTable = excelFile.Worksheets.Add("Отчет");


            mainTable.Cells[5, 0].Column.SetWidth(175,LengthUnit.Pixel);
            mainTable.Cells[5, 1].Column.SetWidth(150,LengthUnit.Pixel);
            mainTable.Cells[0, 2].Column.SetWidth(120,LengthUnit.Pixel);
            mainTable.Cells[5, 3].Column.SetWidth(154,LengthUnit.Pixel);
            mainTable.Cells[5, 4].Column.SetWidth(150,LengthUnit.Pixel);

            mainTable.Cells[0, 2].Value = "ОТЧЕТ ЦЕХА №" + product.FabricId;
            
            
            mainTable.Cells[0, 2].Style.Font.Weight = ExcelFont.MaxWeight;
            
            
            mainTable.Cells[1, 0].Value = "Детали компании \"ООО Компания\" перевезены со склада " + product.StorageName +" № " +  product.StorageNumber;
            mainTable.Cells[2, 0].Value = "Договор № " + product.ContractNumber + " заключен " + product.ContractDate;
            mainTable.Cells[3, 0].Value = "Накладная № " + product.Invoice + " оформлена " + product.InvoiceDate;
            
            
            mainTable.Cells[4, 2].Value = "ПЛАН-ЗАКАЗ";
            mainTable.Cells[4, 2].Style.Font.Weight = ExcelFont.MaxWeight;
            
            mainTable.Cells[5, 0].Value = "Наименование изделия";
            mainTable.Cells[5, 1].Value = "Код";
            mainTable.Cells[5, 2].Value = "Ед. измер.";
            mainTable.Cells[5, 3].Value = "Цена за одну ед.измер";
            mainTable.Cells[5, 4].Value = "Количество";
            
            mainTable.Cells[5, 0].Style.Font.Weight = ExcelFont.MaxWeight;
            mainTable.Cells[5, 1].Style.Font.Weight = ExcelFont.MaxWeight;
            mainTable.Cells[5, 2].Style.Font.Weight = ExcelFont.MaxWeight;
            mainTable.Cells[5, 3].Style.Font.Weight = ExcelFont.MaxWeight;
            mainTable.Cells[5, 4].Style.Font.Weight = ExcelFont.MaxWeight;
            
            mainTable.Cells[6, 0].Value = product.Name;
            mainTable.Cells[6, 1].Value = product.Code;
            mainTable.Cells[6, 2].Value = product.MeasurementUnit;
            mainTable.Cells[6, 3].Value = product.UnitPrice;
            mainTable.Cells[6, 4].Value = product.Amount;
            
            mainTable.Cells[8, 2].Value = "ИНФОРМАЦИЯ";
            mainTable.Cells[8, 2].Style.Font.Weight = ExcelFont.MaxWeight;
            
            mainTable.Cells[9, 0].Value = "ПОСТАВЩИК";
            mainTable.Cells[9, 0].Style.Font.Weight = ExcelFont.MaxWeight;
            
            
            mainTable.Cells[10, 0].Value = "ФАМИЛИЯ";
            mainTable.Cells[11, 0].Value = "ИМЯ";
            mainTable.Cells[12, 0].Value = "НОМЕР ТЕЛЕФОНА";
            mainTable.Cells[13, 0].Value = "ГОРОД";
            
            mainTable.Cells[9, 3].Value = "ПОКУПАТЕЛЬ";
            mainTable.Cells[9, 3].Style.Font.Weight = ExcelFont.MaxWeight;
            
            mainTable.Cells[10, 3].Value = "ФАМИЛИЯ";
            mainTable.Cells[11, 3].Value = "ИМЯ";
            mainTable.Cells[12, 3].Value = "НОМЕР ТЕЛЕФОНА";
            mainTable.Cells[13, 3].Value = "ГОРОД";

            mainTable.Cells[10, 1].Value = product.Diler.Surname;
            mainTable.Cells[11, 1].Value = product.Diler.Name;
            mainTable.Cells[12, 1].Value = product.Diler.PhoneNumber;
            mainTable.Cells[13, 1].Value = product.Diler.City;
            
            mainTable.Cells[10, 4].Value = product.Buyer.Surname;
            mainTable.Cells[11, 4].Value = product.Buyer.Name;
            mainTable.Cells[12, 4].Value = product.Buyer.PhoneNumber;
            mainTable.Cells[13, 4].Value = product.Buyer.City;
            
            
            mainTable.Cells[15, 0].Value = "ОБЩАЯ ЦЕНА ЗАКАЗА:";
            mainTable.Cells[15, 0].Style.Font.Weight = ExcelFont.MaxWeight;
            
            mainTable.Cells[15, 1].Formula = "=D7*E7";
            
            excelFile.Save("Файл.xlsx");

        }
    }
    
    public static class Notice
    {
        public static void SendError(String error)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(error);
            Console.ForegroundColor = ConsoleColor.White;
        }
    }

    public class Product
    {
        public String Name; //Наименование изделия;
        public String Code; //Код изделия
        public String MeasurementUnit; //Единица измерения
        public int UnitPrice; //Цена за единицу измеряемого изделия
        public int FabricId; //Номер цеха
        public int StorageNumber; //Номер склада
        public String StorageName; //Наименование склада
        public int Invoice; //Номер накладной
        public DateTime ContractDate; //Дата в договоре
        public DateTime InvoiceDate; //Дата в накладной
        public int Amount; //Количество
        public int ContractNumber; //Номер договора;
        public Person Diler; //Поставщик
        public Person Buyer; //Покупатель

        public Product()
        {
            this.Diler = new Person();
            this.Buyer = new Person();
        }

        public bool SetParameter(int pNum, String value)
        {
            try
            {
                if (pNum == 0)
                {
                    this.Name = value;
                    return true;
                }
                if (pNum == 1)
                {
                    this.Code = value;
                    return true;
                }
                if (pNum == 2)
                {
                    this.MeasurementUnit = value;
                    return true;
                }
                if (pNum == 3)
                {
                    this.UnitPrice = Convert.ToInt32(value);
                    return CheckNumberSign(this.UnitPrice);
                }
                if (pNum == 4)
                {
                    this.FabricId = Convert.ToInt32(value);
                    return CheckNumberSign(this.FabricId);
                }
                if (pNum == 5)
                {
                    this.StorageNumber = Convert.ToInt32(value);
                    return CheckNumberSign(this.StorageNumber);
                }
                if (pNum == 6)
                {
                    this.StorageName = value;
                    return true;
                }
                if (pNum == 7)
                {
                    this.Invoice = Convert.ToInt32(value);
                    return CheckNumberSign(this.Invoice);
                }
                if (pNum == 8)
                {
                    //ДД.ММ.ГГГГ
                    DateTime? date = ConvertToDate(value);
                    
                    if (date == null)
                    {
                        return false;
                    }

                    this.ContractDate = (DateTime) date;
                    return true;
                }
                if (pNum == 9)
                {
                    //ДД.ММ.ГГГГ
                    DateTime? date = ConvertToDate(value);
                    
                    if (date == null)
                    {
                        return false;
                    }

                    this.InvoiceDate = (DateTime) date;
                    return true;
                }
                if (pNum == 10)
                {
                    this.Amount = Convert.ToInt32(value);
                    return CheckNumberSign(this.Amount);
                }
                if (pNum == 11)
                {
                    this.ContractNumber = Convert.ToInt32(value);
                    return CheckNumberSign(this.ContractNumber);
                }
            }
            catch (FormatException)
            {
                Notice.SendError("Не удалось превести значение к числу, повторите попытку");
                return false;
            }
            
            Console.WriteLine("Product.SetParameter: Введён неизвестный номер параметра (pNum = " + pNum + ")");
            return false;
        }

        

        private static bool CheckNumberSign(int num)
        {
            //num = -12
            if (num <= 0)
            {
                Notice.SendError("Параметр должен быть больше нуля");
                return false;
            }
            else return true;
        }

        private static DateTime? ConvertToDate(String str)
        {
            //a.bb.ccc.dddd
            String[] dateParts = str.Split('.');

            dateParts = new[] {"0","asdw","wdqho"};
            
            if (dateParts.Length != 3)
            {
                Notice.SendError("Неверный формат даты, используйте формат ДД.ММ.ГГГГ");
                return null;
            }

            String dayStr = dateParts[0];
            String mounthStr = dateParts[1];
            String yearStr = dateParts[2];

            if 
            (
                dayStr.Length != 2 ||
                mounthStr.Length != 2 ||
                yearStr.Length != 4
            )
            {
                Notice.SendError("Неверный формат даты, используйте формат ДД.ММ.ГГГГ");
                return null;
            }
                    
                    
            int day;
            int month;
            int year;

            try
            {
                day = Convert.ToInt32(dateParts[0]);
                month = Convert.ToInt32(dateParts[1]);
                year = Convert.ToInt32(dateParts[2]);
            }
            catch (Exception)
            {
                Notice.SendError("Значения должны быть в числовом виде, повторите попытку");
                return null;
            }

            DateTime ret;
            
            try
            {
                ret = new DateTime(year, month, day);
            }
            catch (ArgumentOutOfRangeException)
            {
                Notice.SendError("Указана невозможная дата, попробуйте снова");
                return null;
            }

            return ret;

        }

        
    }

    public class Person
    {
        public String Name;
        public String Surname;
        public String City;
        public String PhoneNumber;

        public Person() {}

        public Person(string name, string surname, string city, string phoneNumber)
        {
            Name = name;
            Surname = surname;
            City = city;
            PhoneNumber = phoneNumber;
        }

        public bool SetParameter(int pNum, String value)
        {
            if (pNum == 0)
            {
                this.Name = value;
                return true;
            }
            if (pNum == 1)
            {
                this.Surname = value;
                return true;
            }
            if (pNum == 2)
            {
                this.City = value;
                return true;
            }
            if (pNum == 3)
            {
                if (value.Length != 11)
                {
                    Notice.SendError("Номер телефона должен быть указан в формате 8XXXXXXXXXX");
                    return false;
                }

                if (value[0] != '8')
                {
                    Notice.SendError("Первая цифра телефонного номера должна быть 8");
                    return false;
                }

                try
                {
                    Convert.ToInt64(value);
                }
                catch (Exception)
                {
                    Notice.SendError("Телефонный номер должен состоять из цифр");
                    return false;
                }
                
                this.PhoneNumber = value;
                return true;
            }
            
            Console.WriteLine("Person.SetParameter: Введён неизвестный номер параметра (pNum = " + pNum + ")");
            return false;
        }
    }
}