using System;
using GemBox.Spreadsheet;

namespace ConsoleApp
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //Создаем переменную, содержащую все параметры продукта внутри себя
            Product product = new Product();
            
            //Создаём первый массив вопросов, позже они в цикле будут направленны пользователю в виде подсказки
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
            
            //Создаем цикл, который проитерирует по всем вопросам и задаст их пользователю
            for (int iterator = 0; iterator < questions.Length; iterator++)
            {
                //Вывод подсказки пользователю с подстановкой вопроса в неё
                Console.Write("Введите " + questions[iterator] + ": ");

                //Читаем ввод от пользователя
                String userInput = Console.ReadLine();
                
                //Вызываем у продукта функцию установки параметра по индексу и результат записываем в переменную статуса 
                bool status = product.SetParameter(iterator, userInput);

                //Если статус установки параметра является неуспешным - уменьшаем переменную итератора
                //это действие заставит цикл повторить вопрос
                if (status == false) iterator--;
            }
            
            //Создаём второй массив вопросов, данные вопросы будут использоваться в двух циклах,
            //а именно в цикле вопросов по поставщику и по покупателю
            String[] personQuestions = new string[]
            {
                "имя",
                "фамилию",
                "город",
                "номер телефона"
            };
            //Создаем цикл вопросов к поставщику
            //Принцип работы данного цикла идентичен первому циклу с основными вопросами
            for (int i = 0; i < personQuestions.Length; i++)
            {
                Console.Write("Введите " + personQuestions[i] + " поставщика: ");
                bool status = product.Diler.SetParameter(i, Console.ReadLine());
                if (status == false) i--;
            }
            //Создаем цикл вопросов к покупателю
            //Принцип работы данного цикла идентичен первому циклу с основными вопросами
            for (int i = 0; i < personQuestions.Length; i++)
            {
                Console.Write("Введите " + personQuestions[i] + " покупателя: ");
                bool status = product.Buyer.SetParameter(i, Console.ReadLine());
                if (status == false) i--;
            }
            
            //Установка параметров в переменную product завершена
            
            
            //Установка лицензии для использование библиотеки ввода и вывода данных в Excel
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            
            //Создание переменной фойла Excel
            ExcelFile excelFile = new ExcelFile();
            
            //Добавление в файл Excell новой таблици с именем Отчет
            ExcelWorksheet mainTable = excelFile.Worksheets.Add("Отчет");
            
            //Запись параметров продукта в таблицу
            WriteProductToTable(mainTable, product);

            //Сохранение файла Excel
            excelFile.Save("Файл.xlsx");

        }

        /// <summary>
        /// Записывает и форматирует в таблицу данные из обекта product
        /// </summary>
        /// <param name="table">таблица</param>
        /// <param name="product">объект с данными</param>
        private static void WriteProductToTable(ExcelWorksheet table, Product product)
        {
            
            //Установка ширины колонок в файле Excel
            table.Cells[0, 0].Column.SetWidth(175,LengthUnit.Pixel);
            table.Cells[0, 1].Column.SetWidth(150,LengthUnit.Pixel);
            table.Cells[0, 2].Column.SetWidth(120,LengthUnit.Pixel);
            table.Cells[0, 3].Column.SetWidth(154,LengthUnit.Pixel);
            table.Cells[0, 4].Column.SetWidth(150,LengthUnit.Pixel);
            
            table.Cells[0, 2].Value = "ОТЧЕТ ЦЕХА №" + product.FabricId;


            table.Cells[0, 2].Style.Font.Weight = ExcelFont.MaxWeight;


            table.Cells[1, 0].Value = "Детали компании \"ООО Компания\" перевезены со склада " + product.StorageName + " № " +
                                      product.StorageNumber;
            table.Cells[2, 0].Value = "Договор № " + product.ContractNumber + " заключен " + product.ContractDate;
            table.Cells[3, 0].Value = "Накладная № " + product.Invoice + " оформлена " + product.InvoiceDate;


            table.Cells[4, 2].Value = "ПЛАН-ЗАКАЗ";
            table.Cells[4, 2].Style.Font.Weight = ExcelFont.MaxWeight;

            table.Cells[5, 0].Value = "Наименование изделия";
            table.Cells[5, 1].Value = "Код";
            table.Cells[5, 2].Value = "Ед. измер.";
            table.Cells[5, 3].Value = "Цена за одну ед.измер";
            table.Cells[5, 4].Value = "Количество";

            table.Cells[5, 0].Style.Font.Weight = ExcelFont.MaxWeight;
            table.Cells[5, 1].Style.Font.Weight = ExcelFont.MaxWeight;
            table.Cells[5, 2].Style.Font.Weight = ExcelFont.MaxWeight;
            table.Cells[5, 3].Style.Font.Weight = ExcelFont.MaxWeight;
            table.Cells[5, 4].Style.Font.Weight = ExcelFont.MaxWeight;

            table.Cells[6, 0].Value = product.Name;
            table.Cells[6, 1].Value = product.Code;
            table.Cells[6, 2].Value = product.MeasurementUnit;
            table.Cells[6, 3].Value = product.UnitPrice;
            table.Cells[6, 4].Value = product.Amount;

            table.Cells[8, 2].Value = "ИНФОРМАЦИЯ";
            table.Cells[8, 2].Style.Font.Weight = ExcelFont.MaxWeight;

            table.Cells[9, 0].Value = "ПОСТАВЩИК";
            table.Cells[9, 0].Style.Font.Weight = ExcelFont.MaxWeight;


            table.Cells[10, 0].Value = "ФАМИЛИЯ";
            table.Cells[11, 0].Value = "ИМЯ";
            table.Cells[12, 0].Value = "НОМЕР ТЕЛЕФОНА";
            table.Cells[13, 0].Value = "ГОРОД";

            table.Cells[9, 3].Value = "ПОКУПАТЕЛЬ";
            table.Cells[9, 3].Style.Font.Weight = ExcelFont.MaxWeight;

            table.Cells[10, 3].Value = "ФАМИЛИЯ";
            table.Cells[11, 3].Value = "ИМЯ";
            table.Cells[12, 3].Value = "НОМЕР ТЕЛЕФОНА";
            table.Cells[13, 3].Value = "ГОРОД";

            table.Cells[10, 1].Value = product.Diler.Surname;
            table.Cells[11, 1].Value = product.Diler.Name;
            table.Cells[12, 1].Value = product.Diler.PhoneNumber;
            table.Cells[13, 1].Value = product.Diler.City;

            table.Cells[10, 4].Value = product.Buyer.Surname;
            table.Cells[11, 4].Value = product.Buyer.Name;
            table.Cells[12, 4].Value = product.Buyer.PhoneNumber;
            table.Cells[13, 4].Value = product.Buyer.City;


            table.Cells[15, 0].Value = "ОБЩАЯ ЦЕНА ЗАКАЗА:";
            table.Cells[15, 0].Style.Font.Weight = ExcelFont.MaxWeight;

            table.Cells[15, 1].Formula = "=D7*E7";
        }
    }
    
    public static class Notice
    {
        /// <summary>
        /// Метод оповещения пользователя об ошибке красным цветом
        /// </summary>
        /// <param name="error">Строка ошибки</param>
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
        public int ContractNumber; //Номер договора
        public Person Diler; //Поставщик
        public Person Buyer; //Покупатель

        public Product()
        {
            this.Diler = new Person();
            this.Buyer = new Person();
        }

        /// <summary>
        /// Устанавливает пераметры продукта по заданному индексу pNum c проверкой
        /// </summary>
        /// <param name="pNum">Индекс параметра</param>
        /// <param name="value">Значение параметра</param>
        /// <returns></returns>
        public bool SetParameter(int pNum, String value)
        {
            //Старт прослушки исклучений, срабатывает если непозможно переконвертировать строку в число 
            try
            {
                switch (pNum)
                {
                    case 0:  //индекс параметра наименование изделия
                        this.Name = value;
                        return true;
                    case 1:  //индекс параметра Код изделия
                        this.Code = value;
                        return true;
                    case 2:  //индекс параметра Единица измерения
                        this.MeasurementUnit = value;
                        return true;
                    case 3:  //индекс параметра Цена за единицу измеряемого изделия
                        this.UnitPrice = Convert.ToInt32(value);
                        return CheckNumberSign(this.UnitPrice);
                    case 4:  //индекс параметра Номер цеха
                        this.FabricId = Convert.ToInt32(value);
                        return CheckNumberSign(this.FabricId);
                    case 5:  //индекс параметра Номер склада
                        this.StorageNumber = Convert.ToInt32(value);
                        return CheckNumberSign(this.StorageNumber);
                    case 6:  //индекс параметра Наименование склада
                        this.StorageName = value;
                        return true;
                    case 7:  //индекс параметра Номер накладной
                        this.Invoice = Convert.ToInt32(value);
                        return CheckNumberSign(this.Invoice);
                    case 8:  //индекс параметра Дата в договоре
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
                    case 9:  //индекс параметра Дата в накладной
                    {
                        //ДД.ММ.ГГГГ
                        DateTime? date = ConvertToDate(value);
                    
                        if (date == null)
                        {
                            return false;
                        }

                        this.InvoiceDate = (DateTime) date;
                        return  true;
                    }
                    case 10: //индекс параметра Количество
                        this.Amount = Convert.ToInt32(value);
                        return CheckNumberSign(this.Amount);
                    case 11: //индекс параметра Номер договора
                        this.ContractNumber = Convert.ToInt32(value);
                        return CheckNumberSign(this.ContractNumber);
                }
            }
            catch (FormatException)
            {
                //Эта часть кода исполняется в случае возникновения исключения в блоке выше
                //в данном случае срабатывает если пользователь ввел вместо числа буквы
                Notice.SendError("Не удалось превести значение к числу, повторите попытку");
                return false;
            }

            return false;
        }
        
        /// <summary>
        /// Функция проверки знака числа с выводом ошибки
        /// </summary>
        /// <param name="num"></param>
        /// <returns></returns>
        private static bool CheckNumberSign(int num)
        {
            if (num <= 0)
            {
                Notice.SendError("Параметр должен быть больше нуля");
                return false;
            }
            else return true;
        }

        /// <summary>
        /// Метод преобразования строки в дату
        /// </summary>
        /// <param name="str">Строка с данными</param>
        /// <returns>Объект даты</returns>
        private static DateTime? ConvertToDate(String str)
        {
            String[] dateParts = str.Split('.');

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
    /// <summary>
    /// Класс персоны
    /// </summary>
    public class Person
    {
        public String Name;
        public String Surname;
        public String City;
        public String PhoneNumber;

        /// <summary>
        /// Пустой констуктор персоны, используется в инициальзации параметров продукта без значений
        /// </summary>
        public Person() {}

        /// <summary>
        /// Констуктор персоны
        /// </summary>
        /// <param name="name">Имя персоны</param>
        /// <param name="surname">Фамилия персоны</param>
        /// <param name="city">Город персоны</param>
        /// <param name="phoneNumber">Номер телефона персоны</param>
        public Person(string name, string surname, string city, string phoneNumber)
        {
            Name = name;
            Surname = surname;
            City = city;
            PhoneNumber = phoneNumber;
        }

        /// <summary>
        /// Устанавливает пераметры персоны по заданному индексу pNum c проверкой
        /// </summary>
        /// <param name="pNum">Индекс параметра персоны</param>
        /// <param name="value">Значения параметра</param>
        /// <returns>Статус установки значения</returns>
        public bool SetParameter(int pNum, String value)
        {
            //Имя персоны
            if (pNum == 0)
            {
                this.Name = value;
                return true;
            }
            //Фамилия персоны
            if (pNum == 1)
            {
                this.Surname = value;
                return true;
            }
            //Город персоны
            if (pNum == 2)
            {
                this.City = value;
                return true;
            }
            //Номер телефона персоны
            if (pNum == 3)
            {
                if (value.Length != 11)
                {
                    Notice.SendError("Номер телефона должен быть указан в формате 8XXXXXXXXXX");
                    return false;
                }

                //Проверяем первую цифру ввода
                if (value[0] != '8')
                {
                    Notice.SendError("Первая цифра телефонного номера должна быть 8");
                    return false;
                }

                //Проходим по всем буквам вводимых данных и проверяем их принадлежность к цифрам
                for (int i = 0; i < value.Length; i++)
                {
                    if (!char.IsDigit(value[i]))
                    {
                        Notice.SendError("Телефонный номер должен состоять из цифр");
                        return false;
                    }
                }

                //Все проверки успешны, устанавливаем значение и возвращаем успешный статус установки значения
                this.PhoneNumber = value;
                return true;
            }
            
            return false;
        }
    }
}