using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace UnitTestScripts
{
    [TestClass]
    public class ConsistanceDiffrencesReportTest
    {
        // основная база тмх
        // СЕ = 88501500, ИИ = 88555400
        List<string> expected = new List<string>()
        {
            "material Двуокись углерода газообразная 1 сорт ГОСТ 8050-85/0/2 asm2 ААА.100.000 Сборка 100/2 м3/1 шт",
            "material Аргон жидкий  высший сорт ГОСТ 10157-2016/0/1 asm2 ААА.100.000 Сборка 100/1 м3/1 шт",
            "material Болт с шестигранной головкой ГОСТ Р ИСО 4014 - М16х65-8.8-A3J/1/2 asm1 ААА.110.000 Сборка 110/1 шт/1 шт asm2 ААА.110.000 Сборка 110/1 шт/2 шт",
            "material Проволока сварочная ДКРХМ 3 БТ ЛК62-0,5 ГОСТ 16130-90/1/2 asm1 ААА.100.000 Сборка 100/1 кг/1 шт asm2 ААА.100.000 Сборка 100/2 кг/1 шт",
            "material Хомут NORMA GBS силовой 26/18/1 W1/1/2 asm1 ААА.100.000 Сборка 100/1 шт/1 шт asm2 ААА.100.000 Сборка 100/2 шт/1 шт",
            "material Болт B.М12-6gх60.88.35.016 ГОСТ 3033-79/1/2 asm1 ААА.100.000 Сборка 100/1 шт/1 шт asm2 ААА.100.000 Сборка 100/2 шт/1 шт",
            "material ААА.110.001 (Деталь 1)/2/4 asm1 ААА.100.000 Сборка 100/1 шт/1 шт asm1 ААА.110.000 Сборка 110/1 шт/1 шт asm2 ААА.100.000 Сборка 100/2 шт/1 шт asm2 ААА.110.000 Сборка 110/1 шт/2 шт",
            "material ААА.120.000 (Комплект поставки)/1/2 asm1 ААА.100.000 Сборка 100/1 шт/1 шт asm2 ААА.100.000 Сборка 100/2 шт/1 шт",
            "material ААА.110.002 (Деталь 2)/0/2 asm2 ААА.110.000 Сборка 110/1 шт/2 шт"
        };
        List<string> actual = new List<string>();

        string testFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\testData.txt";
        [TestInitialize]
        public void Init()
        {
            actual = File.ReadLines(testFile).ToList();
        }

        [TestMethod]
        public void AreEquivalent() => CollectionAssert.AreEquivalent(expected, actual, $"\n\nexpected: \n [{string.Join("\n", expected)}]\n\nactual: \n [{string.Join("\n", actual)}]");

        [TestMethod]
        public void LinesTest()
        {
            for (int i = 0; i < actual.Count; i++)
                Assert.AreEqual(expected[i], actual[i]);
        }

        [TestMethod]
        public void SubstrTest()
        {
            int length = Math.Min(actual.Count, expected.Count);
            for (int i = 0; i < length; i++)
            {
                var actualLineSplited = actual[i].Split(' ', '/');
                var expectedLineSplited = expected[i].Split(' ', '/');

                CollectionAssert.AreEquivalent(expectedLineSplited, actualLineSplited);
            }
        }
    }
}
