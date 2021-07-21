using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace NomenclExcelToJson
{
    class Program
    {
        private static string fileChangesBlankElma = @"c:\ChangesBlankElma.txt";
        private static string fileChangesNomBlank = @"c:\ChangesNomBlank.txt";
        private static string fileIn1CNotFound = @"c:\In1CNotFound.txt";
        private static string fileInBlankNotFound = @"c:\InBlankNotFound.txt";
        private static string fileChangesAll = @"c:\ChangesAll.txt";
        private static string fileElmaWithoutRF = @"c:\elmaWithoutRF.txt";
        private static string fileWrongNames = @"c:\WrongNames.txt";

        static void Main()
        {
            //var currentDirectory = Directory.GetCurrentDirectory();

            var nomFile = File.ReadAllText("C:\\Users\\y.koryukov\\Documents\\ELMA\\Test\\TestBlank\\nom1907.txt");
            var elmaFile = File.ReadAllText("C:\\Users\\y.koryukov\\Documents\\ELMA\\Test\\TestBlank\\elma1907.txt");
            var blankFile = File.ReadAllText("C:\\Users\\y.koryukov\\Documents\\ELMA\\Test\\TestBlank\\blank1907.txt");
            string addressNom = nomFile.Trim().Replace("\t", "").Replace("\\", "/").Replace("\n", " ").Replace("\r", " ");
            string addressElma = elmaFile.Trim().Replace("\t", "").Replace("\\", "/").Replace("\n", " ").Replace("\r", " ");
            string addressBlank = blankFile.Trim().Replace("\t", "").Replace("\\", "/").Replace("\n", " ").Replace("\r", " ");

            var nomenclatureData = GetJsonDataFromFile(addressNom);
            var elmaData = GetJsonDataFromElmaFile(addressElma);
            var blankData = GetJsonDataFromBlankFile(addressBlank);

            //ExchangeFromNomenclature1CAndElma(nomenclatureData, elmaData);
            //ExchangeFromNomenclature1CAndBlank(nomenclatureData, blankData);
            //ExchangeFromElmaAndBlank(elmaData, blankData);
            //ExchangeFromAll(nomenclatureData, elmaData, blankData);
            //var jointObjects = GetJointObjects(nomenclatureData, elmaData, blankData);
            var jointObjects = GetRepeatPositions(elmaData, nomenclatureData, blankData);
            ExportToExcel(jointObjects);
            //GetWrongNames(nomenclatureData);
            //ExportChangesInMultiplicity(jointObjects);

            Console.WriteLine("FINISH!");
            Console.ReadKey();
        }

        private static List<NomObject> GetJsonDataFromFile(string nomFile)
        {
            var jsonDataAll = NomObject.FromJson(nomFile);
            var jsonData = new List<NomObject>();
            foreach (var item in jsonDataAll)
            {
                if (item.RF == "Да")
                {
                    jsonData.Add(item);
                }
            }
            return jsonData;
        }
        private static List<ElmaObject> GetJsonDataFromElmaFile(string elmaFile)
        {
            var elmaDataAll = ElmaObject.FromJsonElma(elmaFile);
            //var elmaData = new List<ElmaObject>();
            //foreach (var item in elmaDataAll)
            //{
            //    if (item.Categories.Contains("РФ"))
            //    {
            //        elmaData.Add(item);
            //    }
            //}
            return elmaDataAll;
        }
        private static List<BlankObject> GetJsonDataFromBlankFile(string blankFile)
        {
            var blankData = BlankObject.FromJsonBlank(blankFile);
            return blankData;
        }

        private static void ExchangeFromNomenclature1CAndElma(List<NomObject> nom1CDatas, List<ElmaObject> elmaObjects)
        {
            foreach (var item in nom1CDatas)
            {
                var elmaItems = elmaObjects.Where(c => c.Guid1C.Trim() == item.GUID1C.Trim());
                if (elmaItems.Any())
                {
                    if (elmaItems.Count() > 1)
                    {
                        // TODO 
                    }
                    else
                    {
                        if (!elmaItems.FirstOrDefault().Categories.Contains("РФ"))
                        {
                            File.AppendAllText(fileElmaWithoutRF, elmaItems.FirstOrDefault().Guid1C + Environment.NewLine);
                        }
                    }
                }
            }
        }

        private static void ExchangeFromElmaAndBlank(List<ElmaObject> elmaObjects, List<BlankObject> blankObjects)
        {
            foreach (var blankItem in blankObjects)
            {
                var elmaItems = elmaObjects.Where(c => c.VendorCode.Trim() == blankItem.VendorCode.Trim());
                if (elmaItems.Any())
                {
                    if (elmaItems.Count() > 1)
                    {
                        foreach (var itm in elmaItems)
                        {
                            File.AppendAllText(fileChangesBlankElma, itm.VendorCode.Trim() + " --- Дубль" + Environment.NewLine);
                        }
                    }
                    foreach (var elmaItem in elmaItems)
                    {
                        if (elmaItem.InPack.Trim() != blankItem.InPack.Trim())
                        {
                            File.AppendAllText(fileChangesBlankElma,
                                $"{elmaItem.VendorCode.Trim()} : в ELMA 'Штук в упаковке' - '{elmaItem.InPack.Trim()}', в бланке заказа - '{blankItem.InPack.Trim()}'" + Environment.NewLine);
                        }
                        if (elmaItem.InRow.Trim() != blankItem.InRow.Trim())
                        {
                            File.AppendAllText(fileChangesBlankElma,
                                $"{elmaItem.VendorCode.Trim()} : в ELMA 'Штук в ряду' - '{elmaItem.InRow.Trim()}', в бланке заказа - '{blankItem.InRow.Trim()}'" + Environment.NewLine);
                        }
                        if (elmaItem.InPallet.Trim() != blankItem.InPallet.Trim())
                        {
                            File.AppendAllText(fileChangesBlankElma,
                                $"{elmaItem.VendorCode.Trim()} : в ELMA 'Штук в поддоне' - '{elmaItem.InPallet.Trim()}', в бланке заказа - '{blankItem.InPallet.Trim()}'" + Environment.NewLine);
                        }
                    }
                }
            }
        }

        private static void ExchangeFromNomenclature1CAndBlank(List<NomObject> nom1CDatas, List<BlankObject> blankObjects)
        {
            var inBlankNotFound = new List<NomObject>();
            var in1CNotfound = new List<BlankObject>();
            foreach (var nom in nom1CDatas)
            {
                var blankItems = blankObjects.Where(c => c.VendorCode.Trim() == nom.VendorCode.Trim());
                if (blankItems.Any())
                {
                    if (blankItems.Count() > 1)
                    {
                        // TODO
                    }
                    foreach (var blankItem in blankItems)
                    {
                        if (nom.InPack.Trim() != blankItem.InPack.Trim())
                        {
                            File.AppendAllText(fileChangesNomBlank,
                                $"{nom.VendorCode.Trim()} - 'Штук в упаковке' {nom.InPack.Trim()}, в бланке заказа {blankItem.InPack.Trim()}" + Environment.NewLine);
                        }
                        if (nom.InRow.Trim() != blankItem.InRow.Trim())
                        {
                            File.AppendAllText(fileChangesNomBlank,
                                $"{nom.VendorCode.Trim()} - 'Штук в ряду' {nom.InRow.Trim()}, в бланке заказа {blankItem.InRow.Trim()}" + Environment.NewLine);
                        }
                        if (nom.InPallet.Trim() != blankItem.InPallet.Trim())
                        {
                            File.AppendAllText(fileChangesNomBlank,
                                $"{nom.VendorCode.Trim()} - 'Штук в поддоне' {nom.InPallet.Trim()}, в бланке заказа {blankItem.InPallet.Trim()}" + Environment.NewLine);
                        }
                        File.AppendAllText(fileChangesNomBlank,
                                $"Из 1С пришло OrderMultiplicity - {nom.OrderMultiplicity.Trim()}" + Environment.NewLine);
                        File.AppendAllText(fileChangesNomBlank, " " + Environment.NewLine);
                    }
                }
                else
                {
                    inBlankNotFound.Add(nom);
                }
            }
            foreach (var blankItem in blankObjects)
            {
                var nomItem = nom1CDatas.Where(c => c.VendorCode.Trim() == blankItem.VendorCode.Trim());
                if (nomItem.Any())
                {

                }
                else
                {
                    in1CNotfound.Add(blankItem);
                }
            }
            foreach (var item in in1CNotfound)
            {
                File.AppendAllText(fileIn1CNotFound, item.VendorCode.Trim() + Environment.NewLine);
            }
            foreach (var item in inBlankNotFound)
            {
                File.AppendAllText(fileInBlankNotFound, item.VendorCode.Trim() + Environment.NewLine);
            }
        }

        private static void ExchangeFromAll(List<NomObject> nom1CDatas, List<ElmaObject> elmaObjects, List<BlankObject> blankObjects)
        {
            foreach (var nom in nom1CDatas)
            {
                var elmaItem = elmaObjects.Where(c => c.VendorCode == nom.VendorCode && c.Guid1C == nom.GUID1C).FirstOrDefault();
                var blankItem = blankObjects.Where(c => c.VendorCode == nom.VendorCode).FirstOrDefault();
                string elmaValue = "";
                string blankValue = "";
                bool isChange = false;
                if (elmaItem != null || blankItem != null)
                {
                    if (elmaItem != null)
                    {
                        if (string.IsNullOrEmpty(elmaItem.InPack))
                        {
                            elmaItem.InPack = "0";
                        }
                        if (string.IsNullOrEmpty(elmaItem.InRow))
                        {
                            elmaItem.InRow = "0";
                        }
                        if (string.IsNullOrEmpty(elmaItem.InPallet))
                        {
                            elmaItem.InPallet = "0";
                        }
                        if (elmaItem.InPack.Trim() != nom.InPack.Trim())
                        {
                            elmaValue += $" В упаковке: 1C - {nom.InPack.Trim()}, elma - {elmaItem.InPack.Trim()}" + Environment.NewLine;
                            isChange = true;
                        }
                        if (elmaItem.InRow.Trim() != nom.InRow.Trim())
                        {
                            elmaValue += $" В строке: 1C - {nom.InRow.Trim()}, elma - {elmaItem.InRow.Trim()}" + Environment.NewLine;
                            isChange = true;
                        }
                        if (elmaItem.InPallet.Trim() != nom.InPallet.Trim())
                        {
                            elmaValue += $" В поддоне: 1C - {nom.InPallet.Trim()}, elma - {elmaItem.InPallet.Trim()}" + Environment.NewLine;
                            isChange = true;
                        }
                    }
                    if (blankItem != null)
                    {
                        if (string.IsNullOrEmpty(blankItem.InPack))
                        {
                            blankItem.InPack = "0";
                        }
                        if (string.IsNullOrEmpty(blankItem.InRow))
                        {
                            blankItem.InRow = "0";
                        }
                        if (string.IsNullOrEmpty(blankItem.InPallet))
                        {
                            blankItem.InPallet = "0";
                        }
                        if (blankItem.InPack.Trim() != nom.InPack.Trim())
                        {
                            blankValue += $" В упаковке: 1C - {nom.InPack.Trim()}, бланк заказа - {blankItem.InPack.Trim()}" + Environment.NewLine;
                            isChange = true;
                        }
                        if (blankItem.InRow.Trim() != nom.InRow.Trim())
                        {
                            blankValue += $" В строке: 1C - {nom.InRow.Trim()}, бланк заказа - {blankItem.InRow.Trim()}" + Environment.NewLine;
                            isChange = true;
                        }
                        if (blankItem.InPallet.Trim() != nom.InPallet.Trim())
                        {
                            blankValue += $" В поддоне: 1C - {nom.InPallet.Trim()}, бланк заказа - {blankItem.InPallet.Trim()}" + Environment.NewLine;
                            isChange = true;
                        }
                    }
                    if (isChange && (!string.IsNullOrEmpty(elmaValue) || !string.IsNullOrEmpty(blankValue)))
                    {
                        File.AppendAllText(fileChangesAll,
                            $"Запись: guid1C - {nom.GUID1C.Trim()}, VendorCode - {nom.VendorCode.Trim()}" + Environment.NewLine);
                        if (!string.IsNullOrEmpty(elmaValue))
                        {
                            File.AppendAllText(fileChangesAll, elmaValue);
                        }
                        if (!string.IsNullOrEmpty(blankValue))
                        {
                            File.AppendAllText(fileChangesAll, blankValue);
                        }
                        File.AppendAllText(fileChangesAll, $"1C multiplicity - {nom.OrderMultiplicity.Trim()}" + Environment.NewLine);
                        File.AppendAllText(fileChangesAll, "-" + Environment.NewLine);
                    }
                }
            }
        }

        private static List<JointObject> GetJointObjects(List<NomObject> nomObjects, List<ElmaObject> elmaObjects, List<BlankObject> blankObjects)
        {
            var jointObjects = new List<JointObject>();
            foreach (var blankObject in blankObjects)
            {
                var jointObject = new JointObject();
                jointObject.BlankVendorCode = blankObject.VendorCode.Trim();
                var elmaObject = elmaObjects.Where(c => c.VendorCode.Trim() == blankObject.VendorCode.Trim()).FirstOrDefault();
                var nomObject = nomObjects.Where(c => c.VendorCode.Trim() == blankObject.VendorCode.Trim()).FirstOrDefault();
                jointObject.NomVendorCode = nomObject?.VendorCode;
                jointObject.BlankInPack = blankObject.InPack;
                jointObject.BlankInPallet = blankObject.InPallet;
                jointObject.BlankInRow = blankObject.InRow;
                jointObject.Code1С = elmaObject?.Code1С;
                jointObject.ElmaInPack = elmaObject?.InPack;
                jointObject.ElmaInPallet = elmaObject?.InPallet;
                jointObject.ElmaInRow = elmaObject?.InRow;
                jointObject.ElmaMultiplicity = elmaObject?.Multiplicity;
                jointObject.GUID1C = elmaObject != null ? elmaObject.Guid1C : nomObject?.GUID1C;
                jointObject.Name = blankObject.Name;
                jointObject.IsWrongName = CheckWrongNames(jointObject.Name).ToString();
                jointObject.NomInPack = nomObject?.InPack;
                jointObject.NomInPallet = nomObject?.InPallet;
                jointObject.NomInRow = nomObject?.InRow;
                jointObject.NomRF = nomObject?.RF;
                jointObject.NomOrderMultiplicity = nomObject?.OrderMultiplicity;
                jointObject.NomMultiplicity = nomObject?.Multiplicity;
                jointObject.RecomMultiplicity = GetNomenclatureValue(nomObject);
                jointObject.BlankRF = "Да";
                jointObjects.Add(jointObject);
                nomObjects.Remove(nomObject);
            }
            foreach (var elem in nomObjects)
            {
                var jointObject = new JointObject();
                var elmaObject = elmaObjects.Where(c => c.VendorCode.Trim() == elem.VendorCode.Trim()).FirstOrDefault();
                jointObject.NomVendorCode = elem.VendorCode.Trim();
                jointObject.Code1С = elem.Code1С;
                jointObject.ElmaInPack = elmaObject?.InPack;
                jointObject.ElmaInPallet = elmaObject?.InPallet;
                jointObject.ElmaInRow = elmaObject?.InRow;
                jointObject.ElmaMultiplicity = elmaObject?.Multiplicity;
                jointObject.GUID1C = elem.GUID1C;
                jointObject.Name = elem.Name;
                jointObject.IsWrongName = CheckWrongNames(jointObject.Name).ToString();
                jointObject.NomInPack = elem.InPack;
                jointObject.NomInPallet = elem.InPallet;
                jointObject.NomInRow = elem.InRow;
                jointObject.NomRF = elem.RF;
                jointObject.NomOrderMultiplicity = elem.OrderMultiplicity;
                jointObject.NomMultiplicity = elem.Multiplicity;
                jointObject.RecomMultiplicity = GetNomenclatureValue(elem);
                jointObjects.Add(jointObject);
            }
            return jointObjects;
        }

        public static void ExportToExcel(List<JointObject> jointObjects)
        {
            var newfileNamesFromJson = "JointObjects12__OOS.xlsx";
            using (MemoryStream stream = new MemoryStream())
            {
                Workbook workbook = new Workbook();
                Style TextStyle = workbook.CreateStyle();
                TextStyle.Number = 49;
                StyleFlag TextFlag = new StyleFlag();
                TextFlag.NumberFormat = true;
                workbook.Worksheets[0].Cells[0, 0].Value = "Наименование в БЗ";
                workbook.Worksheets[0].Cells[0, 1].Value = "Некорректное наименование";
                workbook.Worksheets[0].Cells[0, 2].Value = "Guid1C";
                workbook.Worksheets[0].Cells[0, 3].Value = "Артикул в БЗ";
                workbook.Worksheets[0].Cells[0, 4].Value = "Артикул в 1С";
                workbook.Worksheets[0].Cells[0, 5].Value = "Код в 1С";
                workbook.Worksheets[0].Cells[0, 6].Value = "Единица кратности в 1С";
                workbook.Worksheets[0].Cells[0, 7].Value = "Кратность в 1С";
                workbook.Worksheets[0].Cells[0, 8].Value = "Рекомендуемая кратность";
                workbook.Worksheets[0].Cells[0, 9].Value = "Кратность в ELMA";
                workbook.Worksheets[0].Cells[0, 10].Value = "Категория РФ в БЗ";
                workbook.Worksheets[0].Cells[0, 11].Value = "Категория РФ в 1С";
                workbook.Worksheets[0].Cells[0, 12].Value = "БЗ: в упаковке";
                workbook.Worksheets[0].Cells[0, 13].Value = "1С: в упаковке";
                workbook.Worksheets[0].Cells[0, 14].Value = "ELMA: в упаковке";
                workbook.Worksheets[0].Cells[0, 15].Value = "БЗ: в поддоне";
                workbook.Worksheets[0].Cells[0, 16].Value = "1С: в поддоне";
                workbook.Worksheets[0].Cells[0, 17].Value = "ELMA: в поддоне";
                workbook.Worksheets[0].Cells[0, 18].Value = "БЗ: в ряду";
                workbook.Worksheets[0].Cells[0, 19].Value = "1С: в ряду";
                workbook.Worksheets[0].Cells[0, 20].Value = "ELMA: в ряду";
                workbook.Worksheets[0].Cells[0, 21].Value = "Наименование 1С";
                workbook.Worksheets[0].Cells[0, 22].Value = "Артикул в ELMA";
                var row = 1;
                foreach (var jointObject in jointObjects)
                {
                    Console.WriteLine("Add " + jointObject.BlankVendorCode);
                    workbook.Worksheets[0].Cells[row, 0].Value = jointObject.Name;
                    workbook.Worksheets[0].Cells[row, 1].Value = jointObject.IsWrongName;
                    workbook.Worksheets[0].Cells[row, 2].Value = jointObject.GUID1C;
                    workbook.Worksheets[0].Cells[row, 3].Value = jointObject.BlankVendorCode;
                    workbook.Worksheets[0].Cells[row, 4].Value = jointObject.NomVendorCode;
                    workbook.Worksheets[0].Cells[row, 5].Value = jointObject.Code1С;
                    workbook.Worksheets[0].Cells[row, 6].Value = jointObject.NomOrderMultiplicity;
                    workbook.Worksheets[0].Cells[row, 7].Value = jointObject.NomMultiplicity;
                    workbook.Worksheets[0].Cells[row, 8].Value = jointObject.RecomMultiplicity;
                    workbook.Worksheets[0].Cells[row, 9].Value = jointObject.ElmaMultiplicity;
                    workbook.Worksheets[0].Cells[row, 10].Value = jointObject.BlankRF;
                    workbook.Worksheets[0].Cells[row, 11].Value = jointObject.NomRF;
                    workbook.Worksheets[0].Cells[row, 12].Value = jointObject.BlankInPack;
                    workbook.Worksheets[0].Cells[row, 13].Value = jointObject.NomInPack;
                    workbook.Worksheets[0].Cells[row, 14].Value = jointObject.ElmaInPack;
                    workbook.Worksheets[0].Cells[row, 15].Value = jointObject.BlankInPallet;
                    workbook.Worksheets[0].Cells[row, 16].Value = jointObject.NomInPallet;
                    workbook.Worksheets[0].Cells[row, 17].Value = jointObject.ElmaInPallet;
                    workbook.Worksheets[0].Cells[row, 18].Value = jointObject.BlankInRow;
                    workbook.Worksheets[0].Cells[row, 19].Value = jointObject.NomInRow;
                    workbook.Worksheets[0].Cells[row, 20].Value = jointObject.ElmaInRow;
                    workbook.Worksheets[0].Cells[row, 21].Value = jointObject.Name1C;
                    workbook.Worksheets[0].Cells[row, 22].Value = jointObject.ElmaVendorCode;
                    row++;
                }
                Worksheet sheet;
                sheet = workbook.Worksheets[0];
                sheet.AutoFitColumns();
                string dirPath = "C:\\Users\\y.koryukov\\Documents\\ELMA\\Test\\Blank";
                workbook.Save(dirPath + newfileNamesFromJson, SaveFormat.Xlsx);
                stream.Close();
            }
        }

        private static bool CheckWrongNames(string name)
        {
            if (name.StartsWith("ъ") || name.StartsWith("Ъ"))
            {
                return true;
            }
            return false;
        }

        private static void ExportChangesInMultiplicity(List<JointObject> jointObjects)
        {
            var changedObjects = new List<JointObject>();
            foreach (var jointObject in jointObjects)
            {
                int elmaMultiplicity = -1;
                if (!string.IsNullOrEmpty(jointObject.ElmaMultiplicity))
                {
                    elmaMultiplicity = Int32.Parse(jointObject.ElmaMultiplicity.Trim());
                }
                bool needChange = false;
                int elmaValue = -1;
                int blankValue = -1;
                int nomValue = -1;
                if (!string.IsNullOrEmpty(jointObject.NomOrderMultiplicity))
                {
                    switch (jointObject.NomOrderMultiplicity.Trim())
                    {
                        case "Крт":         // берем данные по ряду
                            if (!string.IsNullOrEmpty(jointObject.ElmaInRow))
                            {
                                elmaValue = Int32.Parse(jointObject.ElmaInRow.Trim());
                            }
                            if (!string.IsNullOrEmpty(jointObject.BlankInRow))
                            {
                                blankValue = Int32.Parse(jointObject.BlankInRow.Trim());
                            }
                            if (!string.IsNullOrEmpty(jointObject.NomInRow))
                            {
                                nomValue = Int32.Parse(jointObject.NomInRow.Trim());
                            }
                            needChange = CheckEquals(elmaMultiplicity, elmaValue, blankValue, nomValue);
                            break;
                        case "Нет":         // берем данные из упаковки
                            if (!string.IsNullOrEmpty(jointObject.ElmaInPack))
                            {
                                elmaValue = Int32.Parse(jointObject.ElmaInPack.Trim());
                            }
                            if (!string.IsNullOrEmpty(jointObject.BlankInPack))
                            {
                                blankValue = Int32.Parse(jointObject.BlankInPack.Trim());
                            }
                            if (!string.IsNullOrEmpty(jointObject.NomInPack))
                            {
                                nomValue = Int32.Parse(jointObject.NomInPack.Trim());
                            }
                            needChange = CheckEquals(elmaMultiplicity, elmaValue, blankValue, nomValue);
                            break;
                        case "под":
                            if (!string.IsNullOrEmpty(jointObject.ElmaInPallet))
                            {
                                elmaValue = Int32.Parse(jointObject.ElmaInPallet.Trim());
                            }
                            if (!string.IsNullOrEmpty(jointObject.BlankInPallet))
                            {
                                blankValue = Int32.Parse(jointObject.BlankInPallet.Trim());
                            }
                            if (!string.IsNullOrEmpty(jointObject.NomInPallet))
                            {
                                nomValue = Int32.Parse(jointObject.NomInPallet.Trim());
                            }
                            needChange = CheckEquals(elmaMultiplicity, elmaValue, blankValue, nomValue);
                            break;
                        case "ряд":
                            if (!string.IsNullOrEmpty(jointObject.ElmaInRow))
                            {
                                elmaValue = Int32.Parse(jointObject.ElmaInRow.Trim());
                            }
                            if (!string.IsNullOrEmpty(jointObject.BlankInRow))
                            {
                                blankValue = Int32.Parse(jointObject.BlankInRow.Trim());
                            }
                            if (!string.IsNullOrEmpty(jointObject.NomInRow))
                            {
                                nomValue = Int32.Parse(jointObject.NomInRow.Trim());
                            }
                            needChange = CheckEquals(elmaMultiplicity, elmaValue, blankValue, nomValue);
                            break;
                        case "упак":
                            if (!string.IsNullOrEmpty(jointObject.ElmaInPack))
                            {
                                elmaValue = Int32.Parse(jointObject.ElmaInPack.Trim());
                            }
                            if (!string.IsNullOrEmpty(jointObject.BlankInPack))
                            {
                                blankValue = Int32.Parse(jointObject.BlankInPack.Trim());
                            }
                            if (!string.IsNullOrEmpty(jointObject.NomInPack))
                            {
                                nomValue = Int32.Parse(jointObject.NomInPack.Trim());
                            }
                            needChange = CheckEquals(elmaMultiplicity, elmaValue, blankValue, nomValue);
                            break;
                        case "шт":         // берем данные из упаковки
                            if (!string.IsNullOrEmpty(jointObject.ElmaInPack))
                            {
                                elmaValue = Int32.Parse(jointObject.ElmaInPack.Trim());
                            }
                            if (!string.IsNullOrEmpty(jointObject.BlankInPack))
                            {
                                blankValue = Int32.Parse(jointObject.BlankInPack.Trim());
                            }
                            if (!string.IsNullOrEmpty(jointObject.NomInPack))
                            {
                                nomValue = Int32.Parse(jointObject.NomInPack.Trim());
                            }
                            needChange = CheckEquals(elmaMultiplicity, elmaValue, blankValue, nomValue);
                            break;
                        default:
                            Console.WriteLine(jointObject.NomOrderMultiplicity + Environment.NewLine);
                            break;
                    }
                }
                if (needChange)
                {
                    changedObjects.Add(jointObject);
                }
            }
            if (changedObjects.Count() > 0)
            {
                ExportToExcel(changedObjects);
            }
        }

        private static bool CheckEquals(int elmaMultiplicity, int elmaValue, int blankValue, int nomValue)
        {
            bool needChange = false;
            if (elmaValue != -1)
            {
                if (blankValue != -1)
                {
                    return needChange = elmaValue != blankValue;
                }
                if (nomValue != -1)
                {
                    return needChange = elmaValue != nomValue;
                }
                if (elmaMultiplicity != -1)
                {
                    return needChange = elmaValue != elmaMultiplicity;
                }
            }
            if (blankValue != -1)
            {
                if (elmaValue != -1)
                {
                    return needChange = blankValue != elmaValue;
                }
                if (nomValue != -1)
                {
                    return needChange = blankValue != nomValue;
                }
                if (elmaMultiplicity != -1)
                {
                    return needChange = blankValue != elmaMultiplicity;
                }
            }
            if (nomValue != -1)
            {
                if (elmaValue != -1)
                {
                    return needChange = nomValue != elmaValue;
                }
                if (blankValue != -1)
                {
                    return needChange = nomValue != blankValue;
                }
                if (elmaMultiplicity != -1)
                {
                    return needChange = nomValue != elmaMultiplicity;
                }
            }
            return needChange;
        }

        private static List<JointObject> GetRepeatPositions(List<ElmaObject> elmaObjects, List<NomObject> nomObjects, List<BlankObject> blankObjects)
        {
            var jointObjects = new List<JointObject>();
            foreach (var blank in blankObjects)
            {
                var joint = new JointObject();
                var nomObject = nomObjects.Where(c => c.VendorCode.Trim() == blank.VendorCode.Trim()).FirstOrDefault();
                if (nomObject == null)
                {
                    Console.WriteLine("nom " + blank.VendorCode + " is null");
                    nomObject = nomObjects.Where(c => c.Name.Trim().Contains(blank.Name.Trim())).FirstOrDefault();
                    if (nomObject == null)
                    {
                        continue;
                    }
                    else
                    {
                        Console.WriteLine("nomName for " + blank.VendorCode + " is not null");
                        joint.Name = blank.Name;
                        joint.Name1C = nomObject.Name;
                        joint.BlankVendorCode = blank.VendorCode;
                        joint.NomVendorCode = nomObject.VendorCode;
                        var elmas = elmaObjects.Where(c => c.Name.Trim().Contains(blank.VendorCode.Trim()));
                        foreach (var item in elmas)
                        {
                            joint.ElmaVendorCode += item.VendorCode.Trim() + " | РФ - " + item.Categories.Contains("РФ") + "     " + Environment.NewLine;
                            joint.GUID1C = item.Guid1C.Trim() + " " + Environment.NewLine;
                        }
                        jointObjects.Add(joint);
                    }
                }

            }
            return jointObjects;
        }

        public static string GetNomenclatureValue(NomObject nomObject)
        {
            string multi = "";
            switch (nomObject.OrderMultiplicity)
            {
                case "под":
                    multi = nomObject.InPallet;
                    break;
                case "ряд":
                    multi = nomObject.InRow;
                    break;
                case "упак":
                    multi = nomObject.InPack;
                    break;
                case "Крт":
                    multi = nomObject.Multiplicity;
                    break;
                case "Шт":
                    multi = nomObject.InPack;
                    break;
                default:
                    multi = "н/а";
                    break;
            }
            return multi;
        }
    }
}