///Скрипт v.2.1 от 28.09.2020 Убраны все static для перехода на IPS 6

using System;
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
namespace EcoDiffReport1
{

    public class Script
    {
        public ICSharpScriptContext ScriptContext { get; private set; }
        public ScriptResult Execute(IUserSession session, ImDocumentData document, Int64[] objectIDs)
        {
            Script1 sc = new Script1();
            sc.Execute(session, document, objectIDs);
            return new ScriptResult(true, document); 
        }
    }

    public class Script1
    {
        public bool Execute(IUserSession session, ImDocumentData document, Int64[] objectIDs)
        {
            //try
            {
                document.Designation = "Сравнение составов";
                if (System.Diagnostics.Debugger.IsAttached)
                    System.Diagnostics.Debugger.Break();
                string file = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
                if (System.IO.File.Exists(file))
                    System.IO.File.Delete(file);
                AddToLog("Запускаем скрипт v 3.0");

                long org_BMZ = 3251710;
                //throw new Exception("jhghjhv");
                standart = MetaDataHelper.GetObjectTypeID("cad00252-306c-11d8-b4e9-00304f19f545"); //Стандартные

                proch = MetaDataHelper.GetObjectTypeID("cad0038d-306c-11d8-b4e9-00304f19f545"); //Прочие

                zagot = MetaDataHelper.GetObjectTypeID("cad001da-306c-11d8-b4e9-00304f19f545"); //Заготовка

                matbase = MetaDataHelper.GetObjectTypeID("cad00170-306c-11d8-b4e9-00304f19f545"); //Материал базовый

                izdelie = MetaDataHelper.GetObjectTypeID("cad00268-306c-11d8-b4e9-00304f19f545" /*Изделия*/);

                sostMaterial = MetaDataHelper.GetObjectTypeID("cad00173-306c-11d8-b4e9-00304f19f545"); //Составной материал

                complects = MetaDataHelper.GetObjectTypeID("cad0025f-306c-11d8-b4e9-00304f19f545"); //Комплекты

                //Типы объектов на которые необходимо проверять заполнено ли количество
                checkKolvoTypes.Add(standart);
                checkKolvoTypes.Add(proch);
                checkKolvoTypes.Add(izdelie);
                checkKolvoTypes.Add(matbase);
                checkKolvoTypes.Add(sostMaterial);

                string ecoObjCaption = String.Empty;
                IDBObject ecoObj = session.GetObject(session.EditingContextID, false);
                if (ecoObj != null)
                {
                    var attr = ecoObj.GetAttributeByGuid(new Guid(Intermech.SystemGUIDs.attributeDesignation), false);
                    if (attr != null)
                        document.Designation = attr.AsString;

                    ecoObjCaption = ecoObj.Caption;
                }

                document.Designation = document.Designation + "  " + DateTime.Now.ToString("dd.MM.yyyy, HH:mm:ss");
                document.DocumentName = document.Designation;
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
                rels.Add(MetaDataHelper.GetRelationTypeID("cad00023-306c-11d8-b4e9-00304f19f545" /*Состоит из*/));
                rels.Add(MetaDataHelper.GetRelationTypeID("cad0019f-306c-11d8-b4e9-00304f19f545" /*Состоит из*/));
                rels.Add(MetaDataHelper.GetRelationTypeID(
                "cad00584-306c-11d8-b4e9-00304f19f545" /*Состав экземпляров и партий изделий*/));

                List<ColumnDescriptor> columns = new List<ColumnDescriptor>();
                int attrId = (Int32)ObligatoryObjectAttributes.CAPTION;
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                ColumnNameMapping.FieldName, SortOrders.NONE, 0));
                attrId = MetaDataHelper.GetAttributeTypeID("cad00267-306c-11d8-b4e9-00304f19f545" /*Количество*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));

                if (addBmzFields)
                {
                    attrId = MetaDataHelper.GetAttributeTypeID("cad005de-306c-11d8-b4e9-00304f19f545" /*Норма расхода*/);
                    columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                    ColumnNameMapping.Guid, SortOrders.NONE, 0));
                }

                attrId = MetaDataHelper.GetAttributeTypeID("cad0038c-306c-11d8-b4e9-00304f19f545" /*Материал*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.ID,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));
                attrId = MetaDataHelper.GetAttributeTypeID("cad005e3-306c-11d8-b4e9-00304f19f545" /*Сортамент*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.ID,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));
                attrId = MetaDataHelper.GetAttributeTypeID("cad00020-306c-11d8-b4e9-00304f19f545" /*Наименование*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));
                if (addBmzFields)
                {
                    attrId = MetaDataHelper.GetAttributeTypeID("84ffec95-9b97-4e83-b7d7-0a19833f171a" /*Организация-источник*/);
                    columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                    ColumnNameMapping.Guid, SortOrders.NONE, 0));

                    //attrId = MetaDataHelper.GetAttributeTypeID( "b1e25726-587e-4bb9-8934-75b85b57d5eb" /*Ответственный за номенклатуру ДМТО (БМЗ)*/);
                    //columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                    //ColumnNameMapping.Guid, SortOrders.NONE, 0));

                    attrId = MetaDataHelper.GetAttributeTypeID("120f681e-048d-4a57-b260-1c3481bb15bc" /*Код АМТО*/);
                    columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                    ColumnNameMapping.Guid, SortOrders.NONE, 0));
                }

                attrId = MetaDataHelper.GetAttributeTypeID(new Guid(Intermech.SystemGUIDs.attributeSubstituteInGroup) /*Номер заменителя в группе*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Relation, ColumnContents.Text,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));

                attrId = MetaDataHelper.GetAttributeTypeID(new Guid(Intermech.SystemGUIDs.attributeManufacturingSign) /*Признак изготовления*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));
                attrId = MetaDataHelper.GetAttributeTypeID(Intermech.SystemGUIDs.attributeDesignation /*Обозначение*/);
                columns.Add(new ColumnDescriptor(attrId, AttributeSourceTypes.Object, ColumnContents.Text,
                ColumnNameMapping.Guid, SortOrders.NONE, 0));

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
                #endregion

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
                Intermech.SystemGUIDs.filtrationBaseVersions, null);

                var ecoItems = dt
                    .Select()
                    .Select(r => new Item())
                    .ToList();

                #endregion

                //сохраняем контекст редактирования
                long sessionid = session.EditingContextID;

                #region Второй состав без извещения
                session.EditingContextID = 0;

                items = new List<ObjInfoItem>();
                items.Add(new ObjInfoItem(headerObjBase.ObjectID));

                dt = DataHelper.GetChildSostavData(items, session, rels, -1, dbrsp, null,
                Intermech.SystemGUIDs.filtrationBaseVersions, null);

                var baseItems = dt
                    .Select()
                    .Select(r => new Item())
                    .ToList();

                #endregion

                //возвращаем контекст
                session.EditingContextID = sessionid;

                #region Сравнение составов



                #endregion

                #region Формирование отчета

                // Заполнение шапки
                AddToLog("fill header");
                DocumentTreeNode headerNode = document.Template.FindNode("Шапка");
                if (headerNode != null)
                {
                    TextData textData = headerNode.FindNode("Извещение") as TextData;
                    if (textData != null)
                    {
                        textData.AssignText(ecoObjCaption, false, false, false);
                    }

                    textData = headerNode.FindNode("ДСЕ") as TextData;
                    if (textData != null)
                    {
                        textData.AssignText(headerObjCaption, false, false, false);
                    }
                }

                DocumentTreeNode docrow = document.Template.FindNode("Строка");
                DocumentTreeNode table = document.FindNode("Таблица");
                foreach (var item in resultComposition)
                {

                    AddToLog("beforecreatenode " + item.ToString());
                }

                int index = 0;
                foreach (var item in resultComposition)
                {
                    index++;
                    // for (int i = 0; i < 30; i++)
                    {
                        if (item.Kolvo1 != item.Kolvo2 || item.HasEmptyKolvo1 != item.HasEmptyKolvo2)
                        {
                            AddToLog("createnode " + item.ToString());
                            DocumentTreeNode node = docrow.CloneFromTemplate(true, true);
                            table.AddChildNode(node, false, false);

                            Write(node, "Индекс", index.ToString());
                            Write(node, "ДМТО", item.MaterialDMTO);

                            Write(node, "Код", item.MaterialCode);
                            Write(node, "Материал", item.MaterialCaption);
                            if (item.Kolvo1 != 0 || !item.HasEmptyKolvo1)
                                Write(node, "Было", Convert.ToString(Math.Round(item.Kolvo1, 4)));
                            else
                                Write(node, "Было", "-");
                            if (item.Kolvo2 != 0 || !item.HasEmptyKolvo2)
                                Write(node, "Будет", Convert.ToString(Math.Round(item.Kolvo2, 4)));
                            else
                                Write(node, "Будет", "-");
                            var diff = Math.Round(item.Kolvo2 - item.Kolvo1, 4);
                            Write(node, "Разница", Convert.ToString(diff)/*Convert.ToString(item.Kolvo2 - item.Kolvo1)*/);
                            Write(node, "ЕдИзм", item.EdIzm);

                            if (item.HasEmptyKolvo1 || item.HasEmptyKolvo2)
                            {
                                (node as RectangleElement).AssignLeftBorderLine(
                                new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                                (node as RectangleElement).AssignRightBorderLine(
                                new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                                (node as RectangleElement).AssignTopBorderLine(
                                new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                                (node as RectangleElement).AssignBottomBorderLine(
                                new BorderLine(Color.Red, BorderStyles.SolidLine, 1), false);
                                AddToLog("item.HasEmptyKolvo  " + item.ToString());
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

                }

                AddToLog("Завершили создание отчета ");
                #endregion
            }
            //catch (Exception ex)
            //{
            //AddToLog(ex.Message);
            //AddToLog(ex.StackTrace);
            //throw ex;
            //}
            return true;
        }


        //Список идентификаторов версий объектов
        public List<long> contextObjects = new List<long>();
        //Список идентификаторов  объектов
        public List<long> contextObjectsID = new List<long>();
        public List<int> checkKolvoTypes = new List<int>();

        public void AddToLog(string text)
        {
            string file = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\script.log";
            text = text + Environment.NewLine;
            System.IO.File.AppendAllText(file, text);
            //AddToOutputView(text);
        }


        private void Write(DocumentTreeNode parent, string tplid, string text)
        {
            TextData td = parent.FindFirstNodeFromTemplate_Recursive(tplid) as TextData;
            if (td != null)
                td.AssignText(text, false, false, false);
        }



        public class Item
        {

            public Item(DataRow row)
            {
                ObjectID = Convert.ToInt64(row["F_OBJECT_ID"]);
                ID = Convert.ToInt64(row["F_PART_ID"]);
                TypeID = Convert.ToInt32(row["F_OBJECT_TYPE"]);
                Caption = row["CAPTION"].ToString();
                ObjectID = Convert.ToInt64(row["F_OBJECT_ID"]);



            }

            public long ID
            {
                get => default;
                private set
                {
                }
            }

            public long ObjectID
            {
                get => default;
                private set
                {
                }
            }

            public string Caption
            {
                get => default;
                private set
                {
                }
            }

            public int TypeID
            {
                get => default;
                private set
                {
                }
            }

            public string CodeAMTO
            {
                get => default;
                private set
                {
                }
            }

            public MeasuredValue Quantity
            {
                get => default;
                private set
                {
                }
            }

            public List<Item> Childs
            {
                get => default;
                private set
                {
                }
            }

            public List<Item> Parents
            {
                get => default;
                set
                {
                }
            }

            public int Property
            {
                get => default;
                set
                {
                }
            }

            public int ManufacutredBy
            {
                get => default;
                private set
                {
                }
            }
        }


    }
}