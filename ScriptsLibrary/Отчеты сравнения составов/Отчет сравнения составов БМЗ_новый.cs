///Скрипт v.3.1 - Добавлены входимости материалов в ближайшие сборки
//Скрипт v.4.0 - Расчет ведется по технологическому составу
using System;
using System.Windows.Forms;
using System.Linq;
using System.Xml;
using Intermech.Interfaces;
using Intermech.Interfaces.Document;
using Intermech.Expert;
using Intermech.Kernel.Search;
using System.Data;
using Intermech.Document.Model;
using System.Collections.Generic;
using Intermech.Interfaces.Compositions;
using System.Collections.Specialized;
using System.Collections;
using System.Drawing;
using Intermech.Expert.Scenarios;
using Intermech.Interfaces.Client;
using Intermech.Navigator.Interfaces.QuickSearch;
using Intermech.Interfaces.Contexts;
using Intermech.Collections;
using Intermech;
using Intermech.Interfaces.Workflow;
namespace EcoDiffReport
{
    public class Script
    {
        public ICSharpScriptContext ScriptContext { get; private set; }

        public ScriptResult Execute(IUserSession session, ImDocumentData document, Int64[] objectIDs)
        {
            Report report = new Report() { /*Устанавливаем режим отчета: true - расширенный, false - обычный*/ compliteReport = true,
             originOrg = "БМЗ"};
            report.Run(session, document, objectIDs);

            if(report.compliteReport)
                document.UpdateLayout(true);

            return new ScriptResult(true, document);
        }
    }

    public class Report
    {
        public bool compliteReport = false;

        public bool writeLog = false;
        public bool writeTestingData = false;

        private string _userError = string.Empty;
        private string _adminError = string.Empty;
        private List<long> _messageRecepients = new List<long>() { 3252010 /*Булычева ЕИ*/, 62180376 /*Бородина ЕВ*/ }; //  51448525 /*Протченко ИВ*/
        private long _asm;
        private long _eco;

        private int _partType;
        private int _complectUnitType;
        private int _asmUnitType;
        private int _complectsType;
        private int _zagotType;
        private int _matbaseType;
        private int _sostMaterialType;
        private int _complexPostType;
        private int _MOType;
        private int _CEType;

        //Список идентификаторов версий объектов
        public List<long> contextObjects = new List<long>();

        //Список идентификаторов  объектов
        public List<long> contextObjectsID = new List<long>();

        public List<int> checkAmountTypes = new List<int>();

        public List<int> enabledTypes = new List<int>();

        public string originOrg;

        /// <summary>
        /// Если не выполняются определенные условия (н-р, не покупное изд), materialid=0, и дальше объект пропускается
        /// </summary>
        /// <param name="row"></param>
        /// <param name="itemsDict">Кортеж словарей в который будут заноситься objectid и Item; linkToObjId и Item. Организует иерархическую связь между Item.</param>
        /// <param name="contextMode"></param>
        /// <returns></returns>
        private Item GetItem(DataRow row, Tuple<Dictionary<long, Item>, Dictionary<long, Item>> itemsDict, bool contextMode)
        {
            Item item = null;

            long linkId = 0;
            object link = row["cad001be-306c-11d8-b4e9-00304f19f545" /*Ссылка на объект*/];
            if (link is long)
                linkId = (long)link;
            long objectId = Convert.ToInt64(row["F_OBJECT_ID"]);

            var itemsDict_byObjId = itemsDict.Item1;
            var itemsDict_byLinkId = itemsDict.Item2;

            if (itemsDict_byObjId.ContainsKey(objectId))
            {
                item = itemsDict_byObjId[objectId];
            }
            else
            {
                var objType = Convert.ToInt32(row["F_OBJECT_TYPE"]);

                if (objType == _sostMaterialType)
                    item = new ComplexMaterialItem();
                else if (objType == _matbaseType || MetaDataHelper.IsObjectTypeChildOf(objType, _matbaseType))
                    item = new MaterialItem();
                else
                    item = new Item();

                item.SourseOrg = Convert.ToString(row["84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/]);
                item.Id = Convert.ToInt64(row["F_PART_ID"]);
                item.ObjectId = Convert.ToInt64(row["F_OBJECT_ID"]);
                item.ObjectType = objType;
                item.LinkToObjId = linkId;
                string designation = Convert.ToString(row[Intermech.SystemGUIDs.attributeDesignation]);
                string name = Convert.ToString(row[Intermech.SystemGUIDs.attributeName]);
                string caption = name;
                if (!string.IsNullOrEmpty(designation))
                    caption = designation + " " + name;
                item.Caption = caption;
                if (!contextMode)
                {
                    if (contextObjects.Contains(item.ObjectId) || contextObjects.Contains(-item.ObjectId)) return null; //Объект входит в состав извещения значит считаем что в составе без извещения его нет
                }
                if (contextObjectsID.Contains(item.Id))
                    item.isContextObject = true;

                //if (item is ComplexMaterialItem)
                //{
                //    var Amount1 = row["b9ab8a1a-e988-4031-9f30-72a95644830a" /*Количество компонента 1 (БМЗ)*/];
                //    if (Amount1 is string)
                //        (item as ComplexMaterialItem).Component1Amount = MeasureHelper.ConvertToMeasuredValue(Convert.ToString(Amount1), false);

                //    var Amount2 = row["8de3b657-868a-47de-8a62-3bbaf7c4994f" /*Количество компонента 2 (БМЗ)*/];
                //    if (Amount2 is string)
                //        (item as ComplexMaterialItem).Component2Amount = MeasureHelper.ConvertToMeasuredValue(Convert.ToString(Amount2), false);

                //    var AmountMain = row["34b64b00-d430-48e4-8692-f52a7485a10a" /*Количество основы (БМЗ)*/];
                //    if (AmountMain is string)
                //        (item as ComplexMaterialItem).MainComponentAmount = MeasureHelper.ConvertToMeasuredValue(Convert.ToString(AmountMain), false);
                //}
                //if (item is MaterialItem)
                //{
                //    var compSostMat = row["54b2e64e-1d4b-4a9d-9633-023f6b6bcd4a" /*Компонент составного материала*/];
                //    if (compSostMat is string)
                //    {
                //        switch (compSostMat.ToString())
                //        {
                //            case "Основа":
                //                (item as MaterialItem).ComplexMaterialComponentValue = MaterialItem.ComplexMaterialComponent.ComponentMain;
                //                break;

                //            case "Компонент 1":
                //                (item as MaterialItem).ComplexMaterialComponentValue = MaterialItem.ComplexMaterialComponent.Component1;
                //                break;

                //            case "Компонент 2":
                //                (item as MaterialItem).ComplexMaterialComponentValue = MaterialItem.ComplexMaterialComponent.Component2;
                //                break;

                //            default:
                //                break;
                //        }
                //    }
                //}
            }

            long parentId = Convert.ToInt64(row["F_PROJ_ID"]);
            Relation lnk = new Relation();
            lnk.Child = item;
            lnk.LinkId = Convert.ToInt64(row["F_PRJLINK_ID"]);
            lnk.RelationTypeId = Convert.ToInt32(row["F_RELATION_TYPE"]);

            // если нашли родительский объект в словаре состава, добавляем связи
            if (itemsDict_byObjId.ContainsKey(parentId))
            {
                Item parent = itemsDict_byObjId[parentId];

                if (parent.ObjectType == _complexPostType)
                    return null;

                parent.RelationsWithChild.Add(lnk);
                lnk.Parent = parent;
                item.RelationsWithParent.Add(lnk);
            }
            object AmountValue = null;
            //для составного не записываем количество, чтобы в расчет составных материалов попадала только се
            if (item.ObjectType != _sostMaterialType)
            {
                // получаем значение количества и записываем в Link
                AmountValue = row["cad00267-306c-11d8-b4e9-00304f19f545"];
                if (AmountValue is string)
                    lnk.Amount = MeasureHelper.ConvertToMeasuredValue(Convert.ToString(AmountValue), false);
            }

            if (lnk.Amount == null && MetaDataHelper.GetAttribute4RelationType(lnk.RelationTypeId,
            MetaDataHelper.GetAttributeTypeID("cad00267-306c-11d8-b4e9-00304f19f545" /*Количество*/)) != null)
            {
                foreach (var checkAmountType in checkAmountTypes)
                {
                    if (lnk.Child != null && (lnk.Child.ObjectType == checkAmountType ||
                    MetaDataHelper.IsObjectTypeChildOf(lnk.Child.ObjectType, checkAmountType)))
                    {
                        lnk.HasEmptyAmount = true;
                    }
                }
            }
            itemsDict_byObjId[objectId] = item;

            //для комплектующих единиц записываем группы замен
            if(item.ObjectType == _complectUnitType)
            {
                object substituteInGroup = row[Intermech.SystemGUIDs.attributeSubstituteInGroup];
                if (substituteInGroup != DBNull.Value)
                {
                    int groupNo = Convert.ToInt32(substituteInGroup);
                    item.isActualReplacement = groupNo == 0;
                    item.isPossableReplacement = groupNo > 0;
                }
                object substitutesGroupNo = row[Intermech.SystemGUIDs.attributeSubstitutesGroupNo];
                if (substitutesGroupNo != DBNull.Value)
                {
                    item.ReplacementGroup = Convert.ToInt32(substitutesGroupNo);
                }
            }

            if (linkId > 0 && item.ObjectType == _complectUnitType)
            {
                itemsDict_byLinkId[linkId] = item;
            }

            object isPocupValue = row["8debd174-928c-4c07-9dc1-423557bea1d7" /*Признак изготовления БМЗ*/];
            if (isPocupValue != DBNull.Value)
            {
                item.isPurchased = Convert.ToInt32(isPocupValue) != 1;
                item.isCoop = Convert.ToInt32(isPocupValue) == 3;
            }

            object matCode = row["120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/];
            if (matCode != DBNull.Value)
            {
                item.MaterialCode = Convert.ToString(matCode);
            }

            long materialId = item.Id;

            if (item.ObjectType == _complectUnitType)
            {
                item.MaterialId = materialId;
            }

            if (item.ObjectType == _complectsType || MetaDataHelper.IsObjectTypeChildOf(item.ObjectType, _complectsType))
            {
                item.MaterialId = materialId;
            }

            if (item.ObjectType != _sostMaterialType && (item.ObjectType == _matbaseType || MetaDataHelper.IsObjectTypeChildOf(item.ObjectType, _matbaseType)))
            {
                item.MaterialId = materialId;
            }

            if (item.ObjectType == _zagotType)
            {
                AmountValue = row["cad005de-306c-11d8-b4e9-00304f19f545" /*Норма расхода*/];
                if (AmountValue is string)
                {
                    lnk.Amount = MeasureHelper.ConvertToMeasuredValue(Convert.ToString(AmountValue), false);
                }

                if (lnk.Amount == null && MetaDataHelper.GetAttribute4RelationType(lnk.RelationTypeId,
                MetaDataHelper.GetAttributeTypeID("cad005de-306c-11d8-b4e9-00304f19f545" /*Норма расхода*/)) != null)
                {
                    lnk.HasEmptyAmount = true;
                }

                object matId = row["cad005e3-306c-11d8-b4e9-00304f19f545" /*Сортамент*/];
                if (matId != DBNull.Value)
                {
                    item.MaterialId = Convert.ToInt64(matId);
                }
            }

            //AddToLog("CreateLink " + lnk.ToString());
            return item;
        }

        private Dictionary<Tuple<long, long>, Item> GetComposition(IDBObject headerObj, DataTable dtCompos, IUserSession session)
        {
            Dictionary<Tuple<long, long>, Item> composition = new Dictionary<Tuple<long, long>, Item>();
            var itemsDict = new Tuple<Dictionary<long, Item>, Dictionary<long, Item>>(new Dictionary<long, Item>(), new Dictionary<long, Item>());

            Item header = new Item();
            header.Id = headerObj.ID;
            header.ObjectId = headerObj.ObjectID;
            header.Caption = headerObj.Caption;
            header.ObjectType = headerObj.ObjectType;
            itemsDict.Item1[header.ObjectId] = header;

            foreach (DataRow row in dtCompos.Rows)
            {
                //if (Convert.ToInt32(row["F_OBJECT_TYPE"]) == _zagotType &&
                //    row["84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/].ToString() != "БМЗ")
                //    continue;

                Item item = GetItem(row, itemsDict, true);
            }

            itemsDict.Item1.Remove(headerObj.ObjectID);

            foreach (var item in itemsDict.Item1.Values)
            {
                //для собираемой единицы
                if (item.ObjectType == _asmUnitType)
                {
                    // указывает на головную сборку
                    if (item.LinkToObjId == _asm)
                        item.RelationsWithParent = new List<Relation>();

                    Item complectUnit = null;
                    //находим комплектующую указывающую на ту же сборку
                    if (itemsDict.Item2.TryGetValue(item.LinkToObjId, out complectUnit))
                    {
                        item.RelationsWithParent = complectUnit.RelationsWithParent;

                        foreach (var relParent in complectUnit.RelationsWithParent)
                        {
                            relParent.Child = item;
                        }
                    }
                }

                //обновляем количество сборок и деталей (констр. сост.) из данных технологического состава
                if (item.ObjectType == _partType || item.ObjectType == _CEType)
                {
                    //(по ссылке на объект не удается связать комплектующую со сборкой или деталью, -- ссылка указывает на ид версии базового объекта, а не текущего)
                    //если комплектующая associatedComplectUnit входит в собираемую с тем же названием, что и item в вышестоящую сборку
                    IEnumerable<string> itemEntersInAsms = item.AmountInAsm.Keys.Select(e => e.Parent.Caption);

                    Item associatedComplectUnit = itemsDict.Item2.Values
                        .Where(e => e.ObjectType == _complectUnitType)
                        .Where(e => e.Caption == item.Caption)
                        .FirstOrDefault(e => e.AmountInAsm.Keys.Select(r => r.Parent.Caption).Where(t => itemEntersInAsms.Contains(t)).Count() > 0);

                    //перезаписываем значения количества для item из данных associatedComplectUnit
                    foreach (Relation item_rwp in item.RelationsWithParent)
                    {
                        if (associatedComplectUnit == null)
                            break;

                        Relation associatedItem_rwp = associatedComplectUnit.RelationsWithParent
                            .FirstOrDefault(e => e.Child.Caption == associatedComplectUnit.Caption);

                        item_rwp.Amount = associatedItem_rwp != null ? associatedItem_rwp.Amount : item_rwp.Amount;
                    }
                }
                //для заготовок находим компл ед, записываем ту же организацию
                if(item.ObjectType == _zagotType)
                {
                    var relMO = item.RelationsWithParent.FirstOrDefault();
                    if(relMO != null)
                    {
                        var relPart = relMO.Parent.RelationsWithParent.FirstOrDefault();
                        if(relPart != null)
                        {
                            Item complectUnit;
                            if(itemsDict.Item2.TryGetValue(relPart.Parent.ObjectId, out complectUnit))
                            {
                                complectUnit.SourseOrg = item.SourseOrg;
                                complectUnit.isCoop = item.isCoop;
                                complectUnit.WriteToReportForcibly = true;
                            }
                        }
                    }
                }
                //if (item is ComplexMaterialItem)
                //{
                //    ComplexMaterialItem complexMaterial = item as ComplexMaterialItem;

                //    foreach (var item_rwc in item.RelationsWithChild)
                //    {
                //        switch ((item_rwc.Child as MaterialItem).ComplexMaterialComponentValue)
                //        {
                //            case MaterialItem.ComplexMaterialComponent.Component1:
                //                item_rwc.Amount = item_rwc.Amount == null ? complexMaterial.Component1Amount : item_rwc.Amount;
                //                break;

                //            case MaterialItem.ComplexMaterialComponent.Component2:
                //                item_rwc.Amount = item_rwc.Amount == null ? complexMaterial.Component2Amount : item_rwc.Amount;
                //                break;

                //            case MaterialItem.ComplexMaterialComponent.ComponentMain:
                //                item_rwc.Amount = item_rwc.Amount == null ? complexMaterial.MainComponentAmount : item_rwc.Amount;
                //                break;

                //            default:
                //                break;
                //        }
                //    }
                //}
            }

            foreach (Item item in itemsDict.Item1.Values)
            {
                Tuple<string, string> exceptionInfo = null;

                if (item.MaterialId == 0)
                    continue;

                bool hasContextObjects = false;
                var itemsAmount = item.GetAmount(true, ref hasContextObjects, ref item.HasEmptyAmount, ref exceptionInfo);

                if (exceptionInfo.Item2.Length > 0)
                {
                    _userError += string.Format("\nДанные объекта {0} (при формировании отчета для сборки {2} по извещению {3}) введены в систему IPS некорректно и были исключены из отчета. Требуется кооректировка данных. За подробностями обращайтесь к администраторам САПР.\n Сообщение:\n{1}\n", exceptionInfo.Item1, exceptionInfo.Item2, session.GetObject(_asm).Caption, session.GetObject(_eco).Caption);

                    _adminError += string.Format("\nУ пользователя {2} при формировании отчета для сборки {3} по извещению {4} возникла ошибка. Данные объекта {0} введены в систему IPS некорректно и были исключены из отчета. Требуется кооректировка данных.\n Сообщение:\n{1}\n", exceptionInfo.Item1, exceptionInfo.Item2, session.UserName, session.GetObject(_asm).Caption, session.GetObject(_eco).Caption);
                }

                Item mainClone = item.Clone();

                if (!hasContextObjects)
                    continue;

                foreach (var itemAmount in itemsAmount)
                {
                    var AmountItemClone = item.Clone();
                    AmountItemClone.HasEmptyAmount = mainClone.HasEmptyAmount;
                    AmountItemClone.AmountSum = itemAmount.Value;

                    //если несколько заготовок используют одинаковый сортамент, то materialId = id_сортамента. cachedItem в этом случае будет сортамент. К сортаменту добавляется количество материала в заготовке использующей этот сортамент.
                    //ключ состоит из ид.ед.изм и materialid, потому что itemsAmount для одного item может быть несколько с разными единицами измерения

                    Item cachedItem;
                    Tuple<long, long> itemKey = new Tuple<long, long>(itemAmount.Key, item.GetKey());
                    if (composition.TryGetValue(itemKey, out cachedItem))
                    {
                        if (cachedItem.AmountSum != null)
                        {
                            cachedItem.AmountSum.Add(AmountItemClone.AmountSum);

                            //обновляем данные по входямостям в сборки
                            foreach (KeyValuePair<Relation, MeasuredValue> kic_AmountInAsm in AmountItemClone.AmountInAsm)
                            {
                                MeasuredValue AmountInAsm = null;
                                if (cachedItem.AmountInAsm.TryGetValue(kic_AmountInAsm.Key, out AmountInAsm))
                                    cachedItem.AmountInAsm[kic_AmountInAsm.Key].Add(AmountInAsm);
                                else
                                    cachedItem.AmountInAsm.Add(kic_AmountInAsm.Key, kic_AmountInAsm.Value);
                            }
                        }
                        else
                            cachedItem.AmountSum = AmountItemClone.AmountSum;
                    }
                    else
                    {
                        composition[itemKey] = AmountItemClone;
                    }
                }
            }
            return composition;
        }

        public bool Run(IUserSession session, ImDocumentData document, Int64[] objectIDs)
        {
            _asm = objectIDs[0];
            _eco = session.EditingContextID;
            document.Designation = "Отчет сравнения составов";

            if (System.Diagnostics.Debugger.IsAttached)
                System.Diagnostics.Debugger.Break();

            string logFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
            if (System.IO.File.Exists(logFile))
                System.IO.File.Delete(logFile);

            string testFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\testData.txt";
            if (System.IO.File.Exists(testFile))
                System.IO.File.Delete(testFile);

            //AddToLog("Запускаем скрипт v 4.0");
            _zagotType = MetaDataHelper.GetObjectTypeID("cad001da-306c-11d8-b4e9-00304f19f545" /*Заготовка*/);
            _asmUnitType = MetaDataHelper.GetObjectTypeID("cad00167-306c-11d8-b4e9-00304f19f545" /*Собираемая единица*/);
            _complectUnitType = MetaDataHelper.GetObjectTypeID("cad00166-306c-11d8-b4e9-00304f19f545" /*Комплектующая единица*/);
            _matbaseType = MetaDataHelper.GetObjectTypeID("cad00170-306c-11d8-b4e9-00304f19f545"/*Материал базовый*/);
            _sostMaterialType = MetaDataHelper.GetObjectTypeID("cad00173-306c-11d8-b4e9-00304f19f545" /*Составной материал*/);
            _complectsType = MetaDataHelper.GetObjectTypeID("cad0025f-306c-11d8-b4e9-00304f19f545"/*Комплекты*/);
            _complexPostType = MetaDataHelper.GetObjectTypeID("c248b4f0-69c6-4521-866a-f9837afcb0b2" /*Комплекты поставки*/);
            _CEType = MetaDataHelper.GetObjectTypeID("cad00132-306c-11d8-b4e9-00304f19f545" /*Сборочные единицы*/);
            _MOType = MetaDataHelper.GetObjectTypeID("cad0016f-306c-11d8-b4e9-00304f19f545" /*Маршрут обработки*/);
            _partType = MetaDataHelper.GetObjectTypeID("cad00250-306c-11d8-b4e9-00304f19f545" /*Деталь*/);
            //Типы объектов на которые необходимо проверять заполнено ли количество
            checkAmountTypes.Add(_complectUnitType);
            checkAmountTypes.Add(_matbaseType);
            checkAmountTypes.Add(_sostMaterialType);

            enabledTypes.Add(_CEType);
            enabledTypes.Add(_partType);
            enabledTypes.Add(_MOType);
            enabledTypes.AddRange(MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00250-306c-11d8-b4e9-00304f19f545" /*Детали*/)));
            enabledTypes.AddRange(MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00185-306c-11d8-b4e9-00304f19f545" /*Техпроцесс базовый*/)));
            enabledTypes.Add(MetaDataHelper.GetObjectTypeID("cad001ff-306c-11d8-b4e9-00304f19f545" /*Цехозаход*/));
            enabledTypes.Add(MetaDataHelper.GetObjectTypeID("cad00178-306c-11d8-b4e9-00304f19f545" /*Операция*/));
            enabledTypes.Add(MetaDataHelper.GetObjectTypeID("cad0017d-306c-11d8-b4e9-00304f19f545" /*Переход*/));
            enabledTypes.Add(MetaDataHelper.GetObjectTypeID("b3ec04e4-9d56-4494-a57b-766d10cdfe27" /*Группа материалов*/));
            enabledTypes.AddRange(MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad00170-306c-11d8-b4e9-00304f19f545" /*Материал базовый*/)));
            enabledTypes.Add(_asmUnitType);
            enabledTypes.Add(_complectUnitType);
            enabledTypes.Add(_zagotType);
            enabledTypes.Add(_sostMaterialType);

            string ecoObjCaption = String.Empty;
            IDBObject ecoObj = session.GetObject(session.EditingContextID, false);
            IDBObject asmObj = session.GetObject(_asm, false);
            if (ecoObj != null)
            {

                var attr = ecoObj.GetAttributeByGuid(new Guid(Intermech.SystemGUIDs.attributeDesignation), false);
                if (attr != null)
                    document.DocumentName = attr.AsString;

                attr = asmObj.GetAttributeByGuid(new Guid(Intermech.SystemGUIDs.attributeDesignation), false);
                if (attr != null)
                    document.DocumentName += "__" + attr.AsString;

                ecoObjCaption = ecoObj.Caption;
            }
            else
            {
                MessageBox.Show("Не выбран контекст редактирования.", "Ошибка формирования отчета сравнения составов (было-стало)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            document.Designation += " " + DateTime.Now.ToString("dd.MM.yyyy, HH:mm:ss");

            //сборка по извещению
            IDBObject headerObj = null;
            //базовая версия
            IDBObject headerObjBase = null;
            var list = session.GetObjectIDVersions(objectIDs[0], false);
            foreach (var item in list)
            {
                IDBObject obj = session.GetObject(item, true);
                if (obj.ModificationID == session.EditingContextID)
                {
                    headerObj = obj;
                }

                if (obj.IsBaseVersion)
                    headerObjBase = obj;
            }

            string headerObjCaption = String.Empty;
            if (headerObj == null)
            {
                //Если нет созданной по извещению берем ту что подали на вход
                headerObj = session.GetObject(objectIDs[0], true);
                headerObjBase = headerObj;

                headerObjCaption = headerObj.Caption;
            }

            if (headerObj == null)
                throw new Exception("Не найдена версия для данного контекста редактирования");
            if (headerObjBase == null)
                throw new Exception("Не найдена базовая версия");
            bool addBmzFields = true;

            #region Инициализация запроса к БД
 
            List<int> rels = new List<int>();

            rels.Add(MetaDataHelper.GetRelationTypeID("cad0019f-306c-11d8-b4e9-00304f19f545" /*Технологический состав*/));
            rels.Add(MetaDataHelper.GetRelationTypeID("cad00023-306c-11d8-b4e9-00304f19f545" /*Состоит из*/));

            List<ColumnDescriptor> columns = new List<ColumnDescriptor>();

            int attrId = (Int32)ObligatoryObjectAttributes.CAPTION;
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
            ColumnNameMapping.FieldName, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID(new Guid(Intermech.SystemGUIDs.attributeSubstituteInGroup) /*Номер заменителя в группе*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID(new Guid(Intermech.SystemGUIDs.attributeSubstitutesGroupNo) /*Номер группы заменителя*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("cad00267-306c-11d8-b4e9-00304f19f545" /*Количество*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("34b64b00-d430-48e4-8692-f52a7485a10a" /*Количество основы (БМЗ)*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("b9ab8a1a-e988-4031-9f30-72a95644830a" /*Количество компонента 1 (БМЗ)*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("8de3b657-868a-47de-8a62-3bbaf7c4994f" /*Количество компонента 2 (БМЗ)*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("54b2e64e-1d4b-4a9d-9633-023f6b6bcd4a" /*Компонент составного материала*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("cad00020-306c-11d8-b4e9-00304f19f545" /*Наименование*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID(Intermech.SystemGUIDs.attributeDesignation /*Обозначение*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("cad001be-306c-11d8-b4e9-00304f19f545" /*Ссылка на объект*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.ID,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("cad005de-306c-11d8-b4e9-00304f19f545" /*Норма расхода*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("cad0038c-306c-11d8-b4e9-00304f19f545" /*Материал*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.ID,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            attrId = MetaDataHelper.GetAttributeTypeID("cad005e3-306c-11d8-b4e9-00304f19f545" /*Сортамент*/);
            columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.ID,
            ColumnNameMapping.Guid, SortOrders.NONE, 0));

            columns.Add(new ColumnDescriptor(
            MetaDataHelper.GetAttributeTypeID("120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/),
            AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Guid, SortOrders.NONE, 0));

            if (addBmzFields)
            {
                attrId = MetaDataHelper.GetAttributeTypeID(new Guid("8debd174-928c-4c07-9dc1-423557bea1d7" /*Признак изготовления БМЗ*/) /*Признак изготовления*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));
            }
            if (addBmzFields)
            {
                attrId = MetaDataHelper.GetAttributeTypeID("84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));
            }

            HybridDictionary tags = new HybridDictionary();
            DBRecordSetParams dbrsp = new DBRecordSetParams(null,
            columns != null
            ? columns.ToArray()
            : null,
            0, null,
            QueryConsts.All);
            dbrsp.Tags = tags;

            List<ObjInfoItem> items = new List<ObjInfoItem>();
            items.Add(new ObjInfoItem(headerObj.ObjectID));

            #endregion Инициализация запроса к БД

            IDBEditingContextsService idbECS = session.GetCustomService(typeof(IDBEditingContextsService)) as IDBEditingContextsService;
            var contextObject = idbECS.GetEditingContextsObject(session.SessionGUID, session.EditingContextID, true, false);
            if (contextObject != null)
            {
                contextObjects = contextObject.GetVersionsID(session.EditingContextID, -1);
                foreach (var objectid in contextObjects)
                {
                    var info = session.GetObjectInfo(objectid);
                    contextObjectsID.Add(info.ID);
                }
            }

            #region Первый состав по извещению

            DataTable dt = DataHelper.GetChildSostavData(items, session, rels, -1, dbrsp, null,
            Intermech.SystemGUIDs.filtrationBaseVersions, null, enabledTypes);

            // Храним пару ид версии объекта + ид. физической величины
            // те объекты у которых посчитали количество
            // состав по извещен
            Dictionary<Tuple<long, long>, Item> ecoComposition = GetComposition(headerObj, dt, session);

            //AddToLog("Первый состав с извещением " + headerObj.ObjectID.ToString());

            #endregion Первый состав по извещению

            //сохраняем контекст редактирования
            long sessionid = session.EditingContextID;
            session.EditingContextID = 0;

            #region Второй состав без извещения

            dt = DataHelper.GetChildSostavData(items, session, rels, -1, dbrsp, null,
            Intermech.SystemGUIDs.filtrationBaseVersions, null, enabledTypes);

            //AddToLog("Второй состав без извещения " + headerObjBase.ObjectID.ToString());
            session.EditingContextID = 0;

            Dictionary<Tuple<long, long>, Item> baseComposition = GetComposition(headerObjBase, dt, session);

            #endregion Второй состав без извещения

            //возвращаем контекст
            session.EditingContextID = sessionid;

            #region Итоговый состав материалов

            List<Material> resultComposition = new List<Material>();

            foreach (var resItem in baseComposition)
            {
                
                Item ecoItem;
                //Находим базовый объект в списке объектов из извещения
                if (ecoComposition.TryGetValue(resItem.Key, out ecoItem))
                {
                    ecoComposition.Remove(resItem.Key);
                }
                else
                {
                    var emptyItem = new Tuple<long, long>(0, resItem.Key.Item2);
                    if (ecoComposition.TryGetValue(emptyItem, out ecoItem))
                    {
                        ecoComposition.Remove(emptyItem);
                    }
                }

                Material material = new Material(resItem.Value, ecoItem);
                resultComposition.Add(material);
            }
            foreach (var ecoItem in ecoComposition)
            {
                Material material = new Material(null, ecoItem.Value);
                resultComposition.Add(material);
            }

            #endregion Итоговый состав материалов

            resultComposition = resultComposition
                .Where(e => e.WriteToReportForcibly || (e.Amount1 != e.Amount2) || e.ReplacementGroupIsChanged || e.ReplacementStatusIsChanged)
                .ToList();

            #region Изменение групп заменителей

            foreach (var item in resultComposition)
            {
                if (item.ReplacementGroupIsChanged || item.ReplacementStatusIsChanged)
                {
                    //с актуальной на допустимую
                    if (item.isActualReplacement1 && item.isPossableReplacement2)
                    {
                        item.SubstractAmount(amountIsIncrease: false);
                        var actual = resultComposition.FirstOrDefault(e => e.isActualReplacement2 && e.ReplacementGroup2 == item.ReplacementGroup2);
                        item.MaterialCaption += actual != null ? string.Format("\nприменяется взамен {0}", actual.Caption) : "";
                    }

                    //с допустимой на актуальную
                    if (item.isPossableReplacement1 && item.isActualReplacement2)
                    {
                        item.SubstractAmount(amountIsIncrease: true);
                        var possable = resultComposition.FirstOrDefault(e => e.isPossableReplacement2 && e.ReplacementGroup2 == item.ReplacementGroup2);
                        item.MaterialCaption += possable != null ? string.Format("\nдопускается замена на {0}", possable.Caption) : "";
                    }

                    //с основной на допустимую
                    if ((!item.isActualReplacement1 && !item.isPossableReplacement1) && item.isPossableReplacement2)
                    {
                        item.SubstractAmount(amountIsIncrease: false);
                        var actual = resultComposition.FirstOrDefault(e => e.isActualReplacement2 && e.ReplacementGroup2 == item.ReplacementGroup2);
                        item.MaterialCaption += actual != null ? string.Format("\nприменяется взамен {0}", actual.Caption) : "";
                    }

                    //с допустимой на основную
                    if (item.isPossableReplacement1 && (!item.isActualReplacement2 && !item.isPossableReplacement2))
                    {
                        item.SubstractAmount(amountIsIncrease: true);
                    }

                    //с основной на актуальную
                    if ((!item.isActualReplacement1 && !item.isPossableReplacement1) && item.isActualReplacement2)
                    {
                        var possable = resultComposition.FirstOrDefault(e => e.isPossableReplacement2 && e.ReplacementGroup2 == item.ReplacementGroup2);
                        item.MaterialCaption += possable != null ? string.Format("\nдопускается замена на {0}", possable.Caption) : "";
                    }

                    //с актуальной на основную
                    if (item.isActualReplacement1 && (!item.isActualReplacement2 && !item.isPossableReplacement2))
                    {
                        continue;
                    }
                }
            }

            #endregion

            List<long> resultCompIds = resultComposition
                .Select(e => e.LinkToObj)
                .Where(e => e > 0)
                .Distinct()
                .ToList();

            resultCompIds.AddRange(
                resultComposition
                .Where(e => e.Type == _zagotType)
                .Select(e => e.MaterialId));


            #region Запись значений атрибутов материалов из соответствующих объектов конструкторского состава

            List<Material> resultComposition_tech = new List<Material>();
            resultComposition_tech.AddRange(resultComposition.Where(e => e.isPurchased || e.isCoop));

            if (resultCompIds.Count > 0)
            {
                IDBObjectCollection col = session.GetObjectCollection(-1);

                List<ConditionStructure> conds = new List<ConditionStructure>();
                conds.Add(new ConditionStructure((Int32)ObligatoryObjectAttributes.F_OBJECT_ID, RelationalOperators.In,
                resultCompIds.ToArray(), LogicalOperators.NONE, 0, false));

                columns = new List<ColumnDescriptor>();

                columns.Add(new ColumnDescriptor((Int32)ObligatoryObjectAttributes.F_ID, AttributeSourceTypes.Auto,
                ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0));

                columns.Add(new ColumnDescriptor((Int32)ObligatoryObjectAttributes.F_OBJECT_ID,
                AttributeSourceTypes.Object,
                ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0));

                columns.Add(new ColumnDescriptor((Int32)ObligatoryObjectAttributes.CAPTION,
                AttributeSourceTypes.Object,
                ColumnContents.Text, ColumnNameMapping.FieldName, SortOrders.NONE, 0));

                columns.Add(new ColumnDescriptor(
                MetaDataHelper.GetAttributeTypeID("120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/),
                AttributeSourceTypes.Object, ColumnContents.Text, ColumnNameMapping.Guid, SortOrders.NONE, 0));

                if (addBmzFields)
                {
                    attrId = MetaDataHelper.GetAttributeTypeID(new Guid("8debd174-928c-4c07-9dc1-423557bea1d7" /*Признак изготовления БМЗ*/));
                    columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                    ColumnNameMapping.Guid, SortOrders.NONE, 0));

                    attrId = MetaDataHelper.GetAttributeTypeID("84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/);
                    columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                    ColumnNameMapping.Guid, SortOrders.NONE, 0));
                }

                tags = new HybridDictionary();
                dbrsp = new DBRecordSetParams(conds.ToArray(),
                columns != null
                ? columns.ToArray()
                : null,
                0, null,
                QueryConsts.All);
                dt = col.Select(dbrsp);

                foreach (DataRow row in dt.Rows)
                {
                    long id = Convert.ToInt64(row["F_OBJECT_ID"]);
                    string name = Convert.ToString(row["CAPTION"]);
                    string code = Convert.ToString(row["120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/]);

                    bool isPurchased = false, isCoop = false;
                    object attrValue = row["8debd174-928c-4c07-9dc1-423557bea1d7" /*Признак изготовления БМЗ*/];
                    if (attrValue is long)
                    {
                        isPurchased = Convert.ToInt32(attrValue) != 1 ? true : false;
                        isCoop = Convert.ToInt32(attrValue) == 3 ? true : false;
                    }

                    string sourseOrg = string.Empty;
                    attrValue = row["84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/];
                    if (attrValue is string)
                        sourseOrg = attrValue.ToString();

                    var materials = resultComposition.Where(x => x.LinkToObj == id || x.MaterialId == id);
                    foreach (var mat in materials)
                    {
                        if (mat != null && (isPurchased || isCoop))
                        {
                            mat.isCoop = isCoop;
                            mat.isPurchased = isPurchased;
                            mat.MaterialCode = code;
                            if(mat.Type == _zagotType)
                                mat.MaterialCaption = name;

                            resultComposition_tech.Add(mat);
                            //AddToLog("mat3 " + mat.ToString());
                        }
                    }
                }
            }

            #endregion Запись значений атрибутов материалов из соответствующих объектов конструкторского состава


            List<Material> reportComp = new List<Material>();

            foreach (var item in resultComposition_tech)
            {
                Material repeatItem = null;
                //когда есть повторение позиции, объединяем с уже записанной
                foreach (var repItem in reportComp)
                {
                    if (repItem.MaterialCaption == item.MaterialCaption && repItem.Type == item.Type && repItem.EdIzm == item.EdIzm)
                        repeatItem = repItem;
                }

                if (repeatItem != null)
                    repeatItem = repeatItem.Combine(item);
                else
                    reportComp.Add(item);
            }

            reportComp.RemoveAll(e => reportComp.Count(i => e.MaterialCaption == i.MaterialCaption && e.EdIzm == i.EdIzm) > 1);

            foreach (var item in reportComp)
            {
                //для объектов из других организаций
                if (item.isCoop || item.SourseOrg != originOrg)
                    item.MaterialCaption += " от " + item.SourseOrg;

                ////оставляем только различающиеся элементы EntersInAsm1 и EntersInAsm2
                //var keys1 = item.EntersInAsm1.Keys.ToList();
                //foreach (var key in keys1)
                //{
                //    Tuple<MeasuredValue, MeasuredValue> value = null;
                //    if (item.EntersInAsm2.TryGetValue(key, out value))
                //    {
                //        if (value.Item1 == null || value.Item2 == null || item.EntersInAsm1[key].Item1 == null || item.EntersInAsm1[key].Item2 == null)
                //        {
                //            item.EntersInAsm1.Remove(key);
                //            item.EntersInAsm2.Remove(key);
                //            continue;
                //        }
                //        if (value.Item1.Value == item.EntersInAsm1[key].Item1.Value && value.Item2.Value == item.EntersInAsm1[key].Item2.Value)
                //        {
                //            item.EntersInAsm1.Remove(key);
                //            item.EntersInAsm2.Remove(key);
                //        }
                //    }
                //}
                item.EntersInAsm1 = item.EntersInAsm1.OrderBy(e => e.Key).ToDictionary(e => e.Key, e => e.Value);
                item.EntersInAsm2 = item.EntersInAsm2.OrderBy(e => e.Key).ToDictionary(e => e.Key, e => e.Value);
            }

            reportComp = reportComp
                .OrderBy(e => e.SourseOrg).ThenBy(e => e.Type).ThenBy(e => e.MaterialCaption)
                .Where(e => (e.Amount1 != e.Amount2)).ToList();

            var coopComp = reportComp.Where(e => e.isCoop || e.SourseOrg != originOrg).ToList();

            reportComp.RemoveAll(e => e.isCoop || e.SourseOrg != originOrg);

            for (int i = 0; i < coopComp.Count; i++)
            {
                if (coopComp[i].Type == _zagotType)
                {
                    Material part = coopComp.FirstOrDefault(e => e.Type == _complectUnitType && e.MaterialId == coopComp[i].MaterialId);
                    int partInd = 0;
                    if (part != null)
                        partInd = coopComp.IndexOf(part);
                    else continue;

                    var temp = coopComp[i-1];
                    coopComp[i-1] = coopComp[partInd];
                    coopComp[partInd] = temp;
                }
            }

            reportComp.AddRange(coopComp);

            if (writeTestingData)
            {
                foreach (var item in reportComp)
                {
                    System.IO.File.AppendAllText(testFile, string.Format("material {0}/{1}/{2}", item.MaterialCaption, item.Amount1, item.Amount2));
                    foreach (var asm in item.EntersInAsm1)
                    {
                        System.IO.File.AppendAllText(testFile, string.Format(" asm1 {0}/{1}/{2}", asm.Key, asm.Value.Item1.ToString(), asm.Value.Item2.ToString()));
                    }
                    foreach (var asm in item.EntersInAsm2)
                    {
                        System.IO.File.AppendAllText(testFile, string.Format(" asm2 {0}/{1}/{2}", asm.Key, asm.Value.Item1.ToString(), asm.Value.Item2.ToString()));
                    }
                    System.IO.File.AppendAllText(testFile, "\n");
                }
            }

            #region Формирование отчета

            // Заполнение шапки
            //AddToLog("fill header");


            ReportWriter reportWriter = new ReportWriter(document, reportComp, ecoObjCaption, headerObjCaption, compliteReport);

            reportWriter.WriteReport();

            //AddToLog("Завершили создание отчета ");

            #endregion Формирование отчета

            if (_adminError.Length > 0)
            {
                IRouterService router = session.GetCustomService(typeof(IRouterService)) as IRouterService;
                router.CreateMessage(session.SessionGUID, _messageRecepients.ToArray(), "Ошибка формирования отчета сравнения составов (было-стало)", _adminError, session.UserID);
            }

            if (_userError.Length > 0)
                MessageBox.Show(_userError, "Ошибка формирования отчета сравнения составов (было-стало)", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            return true;
        }

        public void AddToLog(string text)
        {
            string file = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
            text = text + Environment.NewLine;
            System.IO.File.AppendAllText(file, text);
            //AddToOutputView(text);
        }

        private class ReportWriter
        {
            private ImDocumentData _document;
            private string _ecoCaption;
            private string _asmCaption;
            private List<Material> _resultComposition;
            private bool _compliteReport;

            public ReportWriter(ImDocumentData document, List<Material> resultComposition, string ecoCaption, string asmCaption, bool completeReport)
            {
                _document = document;
                _ecoCaption = ecoCaption;
                _asmCaption = asmCaption;
                _compliteReport = completeReport;
                _resultComposition = resultComposition;
            }

            public void WriteReport()
            {
                DocumentTreeNode headerNode = _document.Template.FindNode("Шапка");
                if (headerNode != null)
                {
                    TextData textData = headerNode.FindNode("Извещение") as TextData;
                    if (textData != null)
                    {
                        textData.AssignText(_ecoCaption, false, false, false);
                    }

                    textData = headerNode.FindNode("ДСЕ") as TextData;
                    if (textData != null)
                    {
                        textData.AssignText(_asmCaption, false, false, false);
                    }
                }

                DocumentTreeNode docrow = _document.Template.FindNode("Строка");
                DocumentTreeNode table = _document.FindNode("Таблица");
                DocumentTreeNode doccol2 = _document.Template.FindNode("Столбец2");
                DocumentTreeNode docrow2 = _document.Template.FindNode("Строка2");

                int N = 0;
                int totalIndex = 0;
                int rowIndex = 0;

                foreach (var item in _resultComposition)
                {
                    //AddToLog("createnode " + item.ToString());
                    DocumentTreeNode node = docrow.CloneFromTemplate(true, true);
                    if (_compliteReport)
                    {

                        #region Запись "было"

                        if (item.EntersInAsm1.Count > 0 || item.EntersInAsm2.Count > 0)
                        {
                            N++;
                            totalIndex++;

                            rowIndex++;
                            table.AddChildNode(node, false, false);

                            Write(node, "Индекс", N.ToString());
                            Write(node, "Код", item.MaterialCode);
                            if(item.MaterialCaption.Length != item.Caption.Length)
                            {
                                Write(node, "Материал", item.MaterialCaption, new CharFormat("arial", 10, CharStyle.Italic));
                            }
                            else Write(node, "Материал", item.MaterialCaption);

                            if (item.Amount1 != 0 || !item.HasEmptyAmount1)
                            {
                                Write(node, "Всего", Math.Round(item.Amount1, 3).ToString() + " " + item.EdIzm);

                                if (N == 1)
                                {
                                    DocumentTreeNode row2 = node.FindNode("Строка2");
                                    if (item.EntersInAsm1.Count != 0)
                                    {
                                        WriteFirstAfterParent(row2, "Куда входит", item.EntersInAsm1.First().Key);
                                        WriteFirstAfterParent(row2, "Количество вхождений", item.EntersInAsm1.First().Value.Item1.Caption);
                                        WriteFirstAfterParent(row2, "Количество сборок", item.EntersInAsm1.First().Value.Item2.Caption);
                                    }

                                    for (int j = 1; j < item.EntersInAsm1.Count; j++)
                                    {
                                        DocumentTreeNode node2 = row2.CloneFromTemplate(true, true);
                                        DocumentTreeNode col2 = node.FindNode("Столбец2");

                                        totalIndex++;
                                        col2.AddChildNode(node2, false, false);
                                        WriteFirstAfterParent(node2, string.Format("Куда входит #{0}", totalIndex), item.EntersInAsm1.Keys.ToList()[j]);
                                        WriteFirstAfterParent(node2, string.Format("Количество вхождений #{0}", totalIndex), item.EntersInAsm1.Values.ToList()[j].Item1.Caption);
                                        WriteFirstAfterParent(node2, string.Format("Количество сборок #{0}", totalIndex), item.EntersInAsm1.Values.ToList()[j].Item2.Caption);
                                    }
                                }

                                if (N > 1)
                                {
                                    DocumentTreeNode row2 = node.FindNode(string.Format("Строка2 #{0}", totalIndex));
                                    if (item.EntersInAsm1.Count != 0)
                                    {
                                        WriteFirstAfterParent(row2, string.Format("Куда входит #{0}", totalIndex), item.EntersInAsm1.First().Key);
                                        WriteFirstAfterParent(row2, string.Format("Количество вхождений #{0}", totalIndex), item.EntersInAsm1.First().Value.Item1.Caption);
                                        WriteFirstAfterParent(row2, string.Format("Количество сборок #{0}", totalIndex), item.EntersInAsm1.First().Value.Item2.Caption);
                                    }

                                    for (int j = 1; j < item.EntersInAsm1.Count; j++)
                                    {
                                        DocumentTreeNode node2 = row2.CloneFromTemplate(true, true);
                                        DocumentTreeNode col2 = node.FindNode(string.Format("Столбец2 #{0}", rowIndex));

                                        totalIndex++;
                                        col2.AddChildNode(node2, false, false);
                                        WriteFirstAfterParent(node2, string.Format("Куда входит #{0}", totalIndex), item.EntersInAsm1.Keys.ToList()[j]);
                                        WriteFirstAfterParent(node2, string.Format("Количество вхождений #{0}", totalIndex), item.EntersInAsm1.Values.ToList()[j].Item1.Caption);
                                        WriteFirstAfterParent(node2, string.Format("Количество сборок #{0}", totalIndex), item.EntersInAsm1.Values.ToList()[j].Item2.Caption);
                                    }
                                }
                            }
                        }

                        #endregion Запись "было"

                        #region Запись "стало"

                        if (item.EntersInAsm1.Count > 0 || item.EntersInAsm2.Count > 0)
                        {
                            totalIndex++;
                            rowIndex++;
                            node = docrow.CloneFromTemplate(true, true);

                            table.AddChildNode(node, false, false);

                            if (item.Amount2 != 0 || !item.HasEmptyAmount2)
                            {
                                Write(node, "Всего", Math.Round(item.Amount2, 3).ToString() + " " + item.EdIzm);

                                DocumentTreeNode row2 = node.FindNode(string.Format("Строка2 #{0}", totalIndex));
                                DocumentTreeNode col2 = node.FindNode(string.Format("Столбец2 #{0}", rowIndex));
                                if (item.EntersInAsm2.Count != 0)
                                {
                                    WriteFirstAfterParent(row2, string.Format("Куда входит #{0}", totalIndex), item.EntersInAsm2.First().Key);
                                    WriteFirstAfterParent(row2, string.Format("Количество вхождений #{0}", totalIndex), item.EntersInAsm2.First().Value.Item1.Caption);
                                    WriteFirstAfterParent(row2, string.Format("Количество сборок #{0}", totalIndex), item.EntersInAsm2.First().Value.Item2.Caption);
                                }

                                for (int j = 1; j < item.EntersInAsm2.Count; j++)
                                {
                                    DocumentTreeNode node2 = row2.CloneFromTemplate(true, true);
                                    col2.AddChildNode(node2, false, false);
                                    totalIndex++;
                                    WriteFirstAfterParent(node2, string.Format("Куда входит #{0}", totalIndex), item.EntersInAsm2.Keys.ToList()[j]);
                                    WriteFirstAfterParent(node2, string.Format("Количество вхождений #{0}", totalIndex), item.EntersInAsm2.Values.ToList()[j].Item1.Caption);
                                    WriteFirstAfterParent(node2, string.Format("Количество сборок #{0}", totalIndex), item.EntersInAsm2.Values.ToList()[j].Item2.Caption);
                                }
                            }
                        }

                        #endregion Запись "стало"
                    }
                    else
                    {
                        N++;
                        table.AddChildNode(node, false, false);

                        Write(node, "Индекс", N.ToString());
                        Write(node, "Код", item.MaterialCode);

                        if (item.MaterialCaption.Length != item.Caption.Length)
                            Write(node, "Материал", item.MaterialCaption, new CharFormat("arial", 10, CharStyle.Italic));
                        else Write(node, "Материал", item.MaterialCaption);

                        if (item.Amount1 != 0 || !item.HasEmptyAmount1)
                            Write(node, "Было", Convert.ToString(Math.Round(item.Amount1, 3)));
                        else
                            Write(node, "Было", "-");
                        if (item.Amount2 != 0 || !item.HasEmptyAmount2)
                            Write(node, "Будет", Convert.ToString(Math.Round(item.Amount2, 3)));
                        else
                            Write(node, "Будет", "-");
                        var diff = Math.Round(item.Amount2 - item.Amount1, 3);
                        Write(node, "Разница", Convert.ToString(diff)/*Convert.ToString(item.Amount2 - item.Amount1)*/);
                        Write(node, "ЕдИзм", item.EdIzm);
                    }

                    if (item.HasEmptyAmount1 || item.HasEmptyAmount2)
                    {
                        (node as RectangleElement).AssignLeftBorderLine(
                        new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                        (node as RectangleElement).AssignRightBorderLine(
                        new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                        (node as RectangleElement).AssignTopBorderLine(
                        new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                        (node as RectangleElement).AssignBottomBorderLine(
                        new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                        //AddToLog("item.HasEmptyAmount  " + item.ToString());
                        foreach (DocumentTreeNode child in node.Nodes)
                        {
                            (child as RectangleElement).AssignLeftBorderLine(
                            new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                            (child as RectangleElement).AssignRightBorderLine(
                            new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                            (child as RectangleElement).AssignTopBorderLine(
                            new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                            (child as RectangleElement).AssignBottomBorderLine(
                            new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                            if (child is TextData)
                            {
                                //AddToLog("child  id = " + child.Id);
                                if ((child as TextData).CharFormat != null)
                                {
                                    CharFormat cf = (child as TextData).CharFormat.Clone();
                                    cf.TextColor = Color.Red;
                                    cf.CharStyle = CharStyle.Bold;
                                    //AddToLog("SetCharFormat");
                                    (child as TextData).SetCharFormat(cf, false, false);
                                }
                                else
                                {
                                    //AddToLog("(child as TextData).CharFormat == null");
                                }
                            }
                        }
                    }
                }
            }

            private DocumentTreeNode AddNode(DocumentTreeNode childNode, string value)
            {
                DocumentTreeNode node = childNode.GetDocTreeRoot().Template.FindNode(value).CloneFromTemplate(true, true);
                childNode.AddChildNode(node, false, false);
                return node;
            }

            private void Write(DocumentTreeNode parent, string tplid, string text)
            {
                TextData td = parent.FindFirstNodeFromTemplate_Recursive(tplid) as TextData;
                if (td != null)
                {
                    td.AssignText(text, false, false, false);
                }
            }
            private void Write(DocumentTreeNode parent, string tplid, string text, CharFormat charFormat = null)
            {
                TextData td = parent.FindFirstNodeFromTemplate_Recursive(tplid) as TextData;
                if (td != null)
                {
                    td.SetCharFormat(charFormat == null ? new CharFormat("arial", 10, CharStyle.Regular) : charFormat, false, false);
                    td.AssignText(text, false, false, false);
                }

            }
            private void WriteFirstAfterParent(DocumentTreeNode parent, string tplid, string text)
            {
                TextData td = parent.FindNode(tplid) as TextData;
                if (td != null)
                    td.AssignText(text, false, false, false);
            }
        }

        private class Material
        {
            public Material(Item baseItem, Item ecoItem)
            {
                var item = ecoItem == null ? baseItem : ecoItem;
                MaterialId = item.MaterialId;
                MaterialCaption = item.Caption;
                Caption = item.Caption;
                Type = item.ObjectType;
                LinkToObj = item.LinkToObjId;
                MaterialCode = item.MaterialCode;
                isPurchased = item.isPurchased;
                isCoop = item.isCoop;
                SourseOrg = item.SourseOrg;
                WriteToReportForcibly = item.WriteToReportForcibly;

                if (baseItem != null)
                {
                    HasEmptyAmount1 = item.HasEmptyAmount;
                    if (baseItem.AmountSum != null)
                    {
                        isActualReplacement1 = baseItem.isActualReplacement;
                        isPossableReplacement1 = baseItem.isPossableReplacement;
                        ReplacementGroup1 = baseItem.ReplacementGroup;

                        Amount1 = baseItem.AmountSum.Value;
                        MeasureId = baseItem.AmountSum.MeasureID;
                        var descr = MeasureHelper.FindDescriptor(MeasureId);

                        EntersInAsm1 = _amountAsmInit(baseItem);

                        if (descr != null)
                            EdIzm = descr.ShortName;

                    }
                    else
                    {
                        HasEmptyAmount1 = true;
                    }
                }

                if (ecoItem != null)
                {
                    HasEmptyAmount2 = ecoItem.HasEmptyAmount;
                    if (ecoItem.AmountSum != null)
                    {
                        isActualReplacement2 = ecoItem.isActualReplacement;
                        isPossableReplacement2 = ecoItem.isPossableReplacement;
                        ReplacementGroup2 = ecoItem.ReplacementGroup;

                        Amount2 = ecoItem.AmountSum.Value;
                        MeasureId = ecoItem.AmountSum.MeasureID;
                        var descr = MeasureHelper.FindDescriptor(MeasureId);

                        EntersInAsm2 = _amountAsmInit(ecoItem);

                        if (descr != null)
                            EdIzm = descr.ShortName;
                    }
                    else
                    {
                        HasEmptyAmount2 = true;
                    }
                }
            }
            private Dictionary<string, Tuple<MeasuredValue, MeasuredValue>> _amountAsmInit(Item item)
            {
                Dictionary<string, Tuple<MeasuredValue, MeasuredValue>> entersInAsmToInit = new Dictionary<string, Tuple<MeasuredValue, MeasuredValue>>();
                foreach (var Amountinasm1 in item.AmountInAsm)
                {
                    var asmAmount = Amountinasm1.Key.Parent.RelationsWithParent.FirstOrDefault() != null ? Amountinasm1.Key.Parent.RelationsWithParent.First().Amount : MeasureHelper.ConvertToMeasuredValue(Convert.ToString("1 шт"), false);
                    var _null = MeasureHelper.ConvertToMeasuredValue(Convert.ToString("0 шт"), false);
                    entersInAsmToInit[Amountinasm1.Key.Parent.Caption] = new Tuple<MeasuredValue, MeasuredValue>(Amountinasm1.Value != null ? Amountinasm1.Value : _null, asmAmount != null ? asmAmount : _null);
                }
                return entersInAsmToInit;
            }
            private string materialCode;
            private string edIzm = string.Empty;

            public bool isPurchased = false;
            public bool isCoop = false;
            public long MaterialId;
            public string MaterialCaption;
            public string Caption = string.Empty;
            public string SourseOrg;
            public int Type;
            public long LinkToObj;
            public bool WriteToReportForcibly = false;
            public bool ReplacementGroupIsChanged
            {
                get
                {
                    return ReplacementGroup1 != ReplacementGroup2 && ReplacementGroup1 > 0;
                }
            }
            public bool ReplacementStatusIsChanged
            {
                get
                {
                    return isActualReplacement1 != isActualReplacement2 || isPossableReplacement1 != isPossableReplacement2;
                }
            }

            public int ReplacementGroup1 = -1;
            public bool isActualReplacement1 = false;
            public bool isPossableReplacement1 = false;
            public int ReplacementGroup2 = -1;
            public bool isActualReplacement2 = false;
            public bool isPossableReplacement2 = false;

            /// <summary>
            /// Код АМТО
            /// </summary>
            public string MaterialCode
            {
                get
                {
                    return materialCode;
                }
                set
                {
                    if (value == null)
                        materialCode = "";
                    else
                        materialCode = value;
                }
            }

            /// <summary>
            /// Количество из базовой версии
            /// </summary>
            public double Amount1 = 0;

            /// <summary>
            /// Количество из версии по извещению
            /// </summary>
            public double Amount2 = 0;

            public long MeasureId;

            public string EdIzm
            {
                get
                {
                    return edIzm;
                }
                set
                {
                    if (value == null)
                        edIzm = "";
                    else
                        edIzm = value;
                }
            }

            public bool HasEmptyAmount1 = false;
            public bool HasEmptyAmount2 = false;

            /// <summary>
            /// Вхождения СЕ (название, кол-во, кол-во се)
            /// </summary>
            public Dictionary<string, Tuple<MeasuredValue, MeasuredValue>> EntersInAsm1 = new Dictionary<string, Tuple<MeasuredValue, MeasuredValue>>();

            public Dictionary<string, Tuple<MeasuredValue, MeasuredValue>> EntersInAsm2 = new Dictionary<string, Tuple<MeasuredValue, MeasuredValue>>();

            public void SubstractAmount(bool amountIsIncrease)
            {
                if (amountIsIncrease)
                {
                    this.Amount1 = 0;

                    foreach (var inAsm1 in this.EntersInAsm1)
                    {
                        Tuple<MeasuredValue, MeasuredValue> value = null;
                        if (this.EntersInAsm2.TryGetValue(inAsm1.Key, out value))
                            inAsm1.Value.Item1.Substract(value.Item1);
                    }
                }
                else
                {
                    this.Amount2 = 0;

                    foreach (var inAsm2 in this.EntersInAsm2)
                    {
                        Tuple<MeasuredValue, MeasuredValue> value = null;
                        if (this.EntersInAsm1.TryGetValue(inAsm2.Key, out value))
                            inAsm2.Value.Item1.Substract(value.Item1);
                    }
                }
            }

            public Material Combine(Material material)
            {
                if (this.MaterialCode.Length > 0)
                    this.MaterialCode = material.MaterialCode;

                this.Amount1 += material.Amount1;

                this.Amount2 += material.Amount2;

                //if (!this.isCoop)
                //    this.SourseOrg = material.SourseOrg;
                    
                foreach (var eia1 in material.EntersInAsm1)
                {
                    if (!this.EntersInAsm1.ContainsKey(eia1.Key))
                        this.EntersInAsm1.Add(eia1.Key, eia1.Value);
                    else
                        this.EntersInAsm1[eia1.Key].Item1.Add(eia1.Value.Item1);
                }

                foreach (var eia2 in material.EntersInAsm2)
                {
                    if (!this.EntersInAsm2.ContainsKey(eia2.Key))
                        this.EntersInAsm2.Add(eia2.Key, eia2.Value);
                    else
                        this.EntersInAsm2[eia2.Key].Item1.Add(eia2.Value.Item1);
                }
                return this;
            }

            public override string ToString()
            {
                return string.Format("MaterialId={0}; MaterialCode={1}; MaterialCaption={2}; Amount1={3}; Amount2={4}",
                MaterialId, MaterialCode, MaterialCaption, Amount1, Amount2);
            }
        }
    }

    /// <summary>
    /// Связь между объектами Item. Через связь записывается исходное значение количества материала
    /// </summary>
    public class Relation
    {
        public long LinkId;
        public int RelationTypeId;
        public Item Child;
        public Item Parent;
        public MeasuredValue Amount;
        public bool HasEmptyAmount = false;

        public IDictionary<long, MeasuredValue> GetAmount(ref bool hasContextObject, ref bool hasemptyAmountRelations, ref Tuple<string, string> exceptionInfo)
        {
            IDictionary<long, MeasuredValue> result = new Dictionary<long, MeasuredValue>();
            IDictionary<long, MeasuredValue> itemsAmount = Parent.GetAmount(false, ref hasContextObject, ref hasemptyAmountRelations, ref exceptionInfo);
            //Количество инициализируется в методе GetItem
            // если значение количества пустое, записываем количество у связи, затем возвращаем
            if (Amount != null)
            {
                if (itemsAmount == null || itemsAmount.Count == 0)
                {
                    result[MeasureHelper.FindDescriptor(Amount).PhysicalQuantityID] = Amount.Clone() as MeasuredValue;
                }
                else
                {
                    try
                    {
                        foreach (var itemAmount in itemsAmount)
                        {
                            itemAmount.Value.Multiply(Amount);
                            result[MeasureHelper.FindDescriptor(Amount).PhysicalQuantityID] = itemAmount.Value;
                        }
                    }
                    catch (Exception ex)
                    {
                        exceptionInfo = new Tuple<string, string>(this.Child.Caption, ex.Message);
                        return null;
                    }
                }
            }
            else
            {
                if (HasEmptyAmount)
                {
                    hasemptyAmountRelations = true;
                    if (Parent == null)
                    {
                        AddToLogForLink("Parent == null" + LinkId);
                    }

                    if (Child == null)
                    {
                        AddToLogForLink("Child == null" + LinkId);
                    }

                    if (Parent != null && Child != null)
                    {
                        string text = "Отсутствует количество. Тип связи " +
                        MetaDataHelper.GetRelationTypeName(RelationTypeId) + " Род. объект " +
                        Parent.ToString() + " Дочерний объект " + Child.ToString();
                        string text1 = "Отсутствует количество. " + " Род. объект '" + Parent.Caption +
                        "' Дочерний объект '" + Child.Caption + "'";

                        //Script1.AddToOutputView(text1);

                        #region Перенесем выполнение кода статического метода в текущее место

                        //IOutputView view = ServicesManager.GetService(typeof(IOutputView)) as IOutputView;
                        //if (view != null)
                        //    view.WriteString("Скрипт сравнения составов", text1);
                        //else
                        //    AddToLogForLink("view == null");

                        #endregion Перенесем выполнение кода статического метода в текущее место

                        AddToLogForLink(text);
                    }
                }

                if (Child.RelationsWithChild == null || Child.RelationsWithChild.Count == 0)
                {
                    //Если это материал и он последний, то его количество равно 0
                    return result; // null
                }

                if (itemsAmount != null)
                {
                    result = itemsAmount;
                }
            }

            return result;
        }

        public override string ToString()
        {
            string res = "Link=" + LinkId.ToString();
            if (Parent != null)
                res = res + " parent= " + Parent.Caption.ToString();
            if (Child != null)
                res = res + " child= " + Child.Caption.ToString();
            res = res + " Количество = " + Convert.ToString(Amount);
            return res;
        }

        public void AddToLogForLink(string text)
        {
            string file = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
            text = text + Environment.NewLine;
            System.IO.File.AppendAllText(file, text);
            //AddToOutputView(text);
        }
    }

    /// <summary>
    /// Объект. Хранит ссылки на связи Link; сумму количества материала (получено из связей)
    /// </summary>
    public class Item
    {
        public override string ToString()
        {
            string objType1 = MetaDataHelper.GetObjectTypeName(ObjectType);
            //return string.Format( " Id={3}; Caption={4}; MaterialId={0}; MaterialCode={1}; MaterialCaption={2}; AmountSum={5}; ObjectType = {6}", MaterialId, MaterialCode, MaterialCaption, Id, Caption, AmountSum, objType1);
            return string.Format(
            " Id={3}; Caption={4}; MaterialId={0}; MaterialCode={1}; MaterialCaption={2}; AmountSum={5}; ObjectType = {6}",
            MaterialId, MaterialCode, Caption, Id, Caption, AmountSum, objType1);
        }

        public virtual Item Clone()
        {
            Item clone = new Item();
            clone.MaterialId = MaterialId;
            clone.MaterialCode = MaterialCode;
            clone.Id = Id;
            clone.ObjectId = ObjectId;
            clone.Caption = Caption;
            clone.ObjectType = ObjectType;
            clone.RelationsWithParent = RelationsWithParent;
            clone.RelationsWithChild = RelationsWithChild;
            clone._AmountInAsm = _AmountInAsm;
            clone.LinkToObjId = LinkToObjId;
            clone.isPurchased = isPurchased;
            clone.isPossableReplacement = isPossableReplacement;
            clone.isActualReplacement = isActualReplacement;
            clone.ReplacementGroup = ReplacementGroup;
            clone.isCoop = isCoop;
            clone.SourseOrg = SourseOrg;
            clone.WriteToReportForcibly = WriteToReportForcibly;
            return clone;
        }
        public bool WriteToReportForcibly = false;
        public string SourseOrg;
        public bool isContextObject = false;
        public bool isPurchased = false;
        public bool isCoop = false;
        public MeasuredValue AmountSum;
        public long MaterialId;
        public string MaterialCode;
        public long Id;
        public long ObjectId;
        public string Caption;
        public long LinkToObjId;
        public int ObjectType;

        //группы замен
        public int ReplacementGroup = -1;
        public bool isActualReplacement = false;
        public bool isPossableReplacement = false;

        /// <summary>
        /// Связь с первым вхождением в сборку; количество элемента из ближайшей связи
        /// </summary>
        public Dictionary<Relation, MeasuredValue> AmountInAsm
        {
            get
            {
                if (_AmountInAsm.Keys.Count == 0)
                {
                    foreach (var rel in RelationsWithParent)
                    {
                        if (rel.Parent.ObjectType == MetaDataHelper.GetObjectType(new Guid("cad00167-306c-11d8-b4e9-00304f19f545" /*Собираемая единица*/)).ObjectTypeID)
                        {
                            _AmountInAsm.Add(rel, rel.Amount);
                            return _AmountInAsm;
                        }

                        if (rel.Parent.ObjectType == MetaDataHelper.GetObjectType(new Guid(SystemGUIDs.objtypeAssemblyUnit)).ObjectTypeID)
                            _AmountInAsm.Add(rel, rel.Amount);
                        else
                        {
                            var nextAsms = rel.Parent.AmountInAsm;
                            foreach (var na in nextAsms)
                            {
                                _AmountInAsm[na.Key] = rel.Amount;
                            }
                        }
                    }
                }
                return _AmountInAsm;
            }
        }

        //   public MeasuredValue Amount;
        /// <summary>
        /// Связи с дочерними объектами
        /// </summary>
        public List<Relation> RelationsWithChild = new List<Relation>();

        /// <summary>
        /// Связи с родительскими объектами
        /// </summary>
        public List<Relation> RelationsWithParent = new List<Relation>();

        public bool HasEmptyAmount = false;

        private Dictionary<Relation, MeasuredValue> _AmountInAsm = new Dictionary<Relation, MeasuredValue>();

        public IDictionary<long, MeasuredValue> GetAmount(bool checkContextObject, ref bool hasContextObject, ref bool hasemptyAmountRelations, ref Tuple<string, string> exceptionInfo)
        {
            MeasuredValue measuredValue = null;
            IDictionary<long, MeasuredValue> result = new Dictionary<long, MeasuredValue>();
            if (this.isContextObject) hasContextObject = true;

            if (exceptionInfo == null)
                exceptionInfo = new Tuple<string, string>(this.Caption, string.Empty);

            foreach (Relation relation in RelationsWithParent)
            {
                bool hasContextObject1 = true;
                if (checkContextObject)
                    hasContextObject1 = false;
                IDictionary<long, MeasuredValue> itemsAmount = relation.GetAmount(ref hasContextObject1, ref hasemptyAmountRelations, ref exceptionInfo);

                hasContextObject = hasContextObject | hasContextObject1;
                if (checkContextObject && !hasContextObject1 && !this.isContextObject) continue;
                if (itemsAmount == null)
                {
                    continue;
                }

                foreach (var itemAmount in itemsAmount)
                {
                    if (!result.TryGetValue(itemAmount.Key, out measuredValue))
                    {
                        result[itemAmount.Key] = itemAmount.Value.Clone() as MeasuredValue;
                        continue;
                    }

                    try
                    {
                        measuredValue.Add(itemAmount.Value);
                    }
                    catch (Exception ex)
                    {
                        exceptionInfo = new Tuple<string, string>(this.Caption, ex.Message);
                        return null;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Если MaterialId == 0 вернет ObjectId, в противном случае -- MaterialId.
        /// </summary>
        /// <returns></returns>
        public long GetKey()
        {
            return this.MaterialId != 0
            ? this.MaterialId
            : this.ObjectId;
        }
    }

    /// <summary>
    /// Составной материал
    /// </summary>
    public class ComplexMaterialItem : Item
    {
        public MeasuredValue MainComponentAmount;
        public MeasuredValue Component1Amount;
        public MeasuredValue Component2Amount;
    }

    public class MaterialItem : Item
    {
        public ComplexMaterialComponent ComplexMaterialComponentValue = ComplexMaterialComponent.Empty;

        public enum ComplexMaterialComponent
        {
            Component1 = 1,
            Component2 = 2,
            ComponentMain = 3,
            Empty = 4
        };
    }
}