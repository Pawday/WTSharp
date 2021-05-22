using System;
using System.Linq;

namespace EnumsAndArrays
{
    
    enum TheDeclaredEnum
    {
        DefaultIndex,
        SpecifiedIndex = 10
    }

    enum SpaceAreaRoomType //перечисление типа помещения по заданию
    {
        Commercial,
        Living,
        Private,
        Office,
        Medicine
    }

    internal class Program
    {
        public static void Main(string[] args)
        {
            //Задание 1
            Console.WriteLine("Задание 1");
            TheDeclaredEnum theEnum;

            theEnum = TheDeclaredEnum.DefaultIndex;
            Console.WriteLine($"{theEnum}: {(int)theEnum}");
            theEnum += 10;
            Console.WriteLine($"{theEnum}: {(int)theEnum}");

            
            //Задание 2 а)
            //Объявление двумерного массиваа массивов
            //с данними о машинах и площадей их дверей
            Console.WriteLine("Задание 2.1");
            var cars = new int[2][];

            for (int i = 0; i < cars.Length; i++)
            {
                cars[i] = new int[4];
                for (int j = 0; j < cars[i].Length; j++)
                {
                    cars[i][j] = Int32.Parse(Console.ReadLine());
                }
            }
            
            //Задание 2 б)
            //Объявить одномерный массив символов
            //и посчитать колличество букв а
            Console.WriteLine("Задание 2.2");
            var chars = new char[]{'С','П','Б','Г','Э','У'};
            Console.WriteLine("Колличество букв \"а\" в массиве chars = " + chars.Count((c) =>  c == 'a' ));
            
            //Задание 2 в)
            //объявить двумерный массив действитеьлных чисел,
            //заполнить его и отнять номер столбца из каждого числа
            var floats = new float[5][];
            for (int i = 0; i < floats.Length; i++)
            {
                floats[i] = new float[5];
                for (int j = 0; j < floats[i].Length; j++)
                {
                    Console.WriteLine("Введите рациональное число");
                    floats[i][j] = Convert.ToSingle(Console.ReadLine());
                    floats[i][j] -= j;
                }
            }
            
            
            //Задание 3

            int brs_num = 8;
            float a_dmi = brs_num - 10.5f;
            float b_dmi = brs_num * 3;
            float x_dmi = brs_num + 0.35f;

            float result = F(a_dmi, b_dmi, x_dmi);
            
            //Результат данной провернки всегда является истиной
            //так как результат побитового умножения
            //любого числа на 0b100 (4) всегда чётен
            
            if ((brs_num & 0b100) % 2 == 0)
                Console.WriteLine((int) result);
            else
                Console.WriteLine(result);
            
        }
        static float F(float a, float b, float x)
        {
            if (x < 0 && b != 0)
                return a * x * x + b;

            if (x > 0 && b == 0)
                return (x - a) / (x - b);

            return x / b;
        }
    }
}