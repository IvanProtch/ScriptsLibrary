//Выгрузка PDF из констр ИИ и СБ
using System;
using System.IO;
using System.Data;
using System.Linq;
using System.Xml;
using System.Text;
using System.ComponentModel.Design;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using Intermech;
using Intermech.Interfaces;
using Intermech.Interfaces.Client;
using Intermech.Interfaces.Compositions;
using Intermech.Kernel.Search;
using Intermech.Kernel;
using Intermech.Client.Core;
using Intermech.Interfaces.Workflow;
using System.Diagnostics;

public class Script
{
    public ICSharpScriptContext ScriptContext { get; set; }
    
    const string filePath = @"D:\ВЫГРУЗКА из IPS\";
    int lvl = -1; // Количество уровней для обработки(-1 = выгрузка полного состава)
    List<int> RelID = new List<int> { 1004 }; // id связи "документация на изделие"
    List<int> RelSostID = new List<int> { 1, 1007 };
    
    bool loadIerarchly = false; // выгружать иерархию
    bool createNonIerarchedDir = true; // создать допольнительную папку со всеми файлами без иерархии. Работает только если loadIerarchly = true
    
    public void Execute(IActivity activity)
    {
        if (Debugger.IsAttached)
            Debugger.Break();
        IUserSession session = activity.Session;
        
        for (int i = 0; i < activity.Attachments.Count; i++)
        {
            IDBObject attachment = activity.Attachments[i].Object;
            string fullPath = Path.Combine(filePath, attachment.NameInMessages);
            
            if (!Directory.Exists(fullPath))
            {
                Directory.CreateDirectory(fullPath);
            }
            
            // проходим по вложениям
            if (activity.Attachments[i].ObjectType == 1074 || activity.Attachments[i].ObjectType == 1078 || activity.Attachments[i].ObjectType ==  2760) // сб и конст. ии
            {
                // выгружаем файлы сборки
                DataTable dt = LoadItemIds(session, activity.Attachments[i].ObjectID, RelID, MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad0057f-306c-11d8-b4e9-00304f19f545")/*Конструкторсие документы*/), 1, true);
                // выгружаем файлы ИИ
                if(activity.Attachments[i].ObjectType ==  2760)
                    dt = LoadItemIds(session, activity.Attachments[i].ObjectID, RelSostID, MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad0057f-306c-11d8-b4e9-00304f19f545")/*Конструкторсие документы*/), 2, true);
                if (dt != null)
                {
                    foreach (DataRow rw in dt.Rows)
                    {
                        IDBObject objDocSB = session.GetObject(Convert.ToInt64(rw[0]));
                        string nextFilePath = string.Format(filePath + objDocSB.Caption);
                        //DirectoryInfo di1 = Directory.CreateDirectory(nextFilePath);
                        WriteFile(session, objDocSB, attachment,  fullPath, loadIerarchly);
                        
                        if (loadIerarchly && createNonIerarchedDir)
                            WriteFile(session, objDocSB, attachment,  fullPath, false);
                    }
                }
                // получаем состав
                DataTable objSostIDs = LoadItemIds(session, attachment.ObjectID, RelSostID, null, lvl, true);
                foreach (DataRow objSostID in objSostIDs.Rows)
                {
                    // выгружаем файлы состава
                    IDBObject objSost = session.GetObject(Convert.ToInt64(objSostID[0]));
                    DataTable dtObjSost = LoadItemIds(session, objSost.ObjectID, RelID, MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid("cad0057f-306c-11d8-b4e9-00304f19f545")/*Конструкторсие документы*/), 1, true);
                    if (dtObjSost != null)
                    {
                        foreach (DataRow rwDocSost in dtObjSost.Rows)
                        {
                            IDBObject objDocSost = session.GetObject(Convert.ToInt64(rwDocSost[0]));
                            WriteFile(session, objDocSost, objSost,   fullPath, loadIerarchly);
                            
                            if (loadIerarchly && createNonIerarchedDir)
                                WriteFile(session, objDocSost, objSost,  fullPath, false);
                        }
                    }
                }
            }
        }
    }
    
    // добавление файла к объекту
    private void WriteFile(IUserSession session, IDBObject doc, IDBObject objSost, string path, bool withIerarchly)
    {
        IDBAttribute fileAttr = doc.Attributes.AddAttribute(MetaDataHelper.GetAttributeTypeID(new Guid(Intermech.SystemGUIDs.attributeFile)), false);
        
        for (int i = 0; i < fileAttr.ValuesCount; i++)
        {
            MemoryStream ms = new MemoryStream();
            BlobProcReader bpr = new BlobProcReader(doc.ObjectID, AttributableElements.Object, MetaDataHelper.GetAttributeTypeID(new Guid(Intermech.SystemGUIDs.attributeFile)), i, 0, ms, null, null);
            bpr.ReadData();
            BlobInformation bi = bpr.BlobInformation;
            
            if (bi.FileType == FileTypes.ftNormal || bi.FileType == FileTypes.ftAuthentical)
            {
                if (bi.FileName.Contains(".pdf") || bi.FileName.Contains(".PDF") || bi.FileName.Contains(".tif") || bi.FileName.Contains(".TIF"))
                {
                    string fileAttrValue = fileAttr.Values[i].ToString();
                    string expPath = string.Empty;
                    string filePath = string.Empty;


                    if (fileAttrValue != "")
                    {
                        IDBAttribute izmN = objSost.GetAttributeByGuid(new Guid("cad00770-306c-11d8-b4e9-00304f19f545" /*Номер изменения*/));
                        string fileName = Path.GetFileName(fileAttrValue);
                        if (izmN != null)
                            fileName = fileName.Insert(fileName.IndexOf(".pdf"), "_№изм."+((izmN.Value.ToString() == "") ? "0" : izmN.Value.ToString()));

                        if (withIerarchly)
                        {
                            if (fileName.Length != fileAttrValue.Length)
                            {
                                expPath = fileAttrValue.Remove(fileAttrValue.LastIndexOf("\\"));
                            }
                        }
                        else
                        {
                            expPath = "Все файлы";
                        }
                        
                        string dirPath = Path.Combine(path, expPath);
                        if (!Directory.Exists(dirPath))
                            Directory.CreateDirectory(dirPath);
                        
                        filePath = Path.Combine(path, expPath, fileName);
                        
                        // добавляем
                        using (FileStream fs = new FileStream(filePath, FileMode.Create))
                        {
                            BlobProcReader bpReader = new BlobProcReader(doc.ObjectID, AttributableElements.Object, MetaDataHelper.GetAttributeTypeID(new Guid(Intermech.SystemGUIDs.attributeFile)), i, 0, fs, null, null);
                            bpReader.ReadData();
                            fs.Flush();
                        }
                    }
                }
            }
        }
    }
    
    // получаем состав/применяемость
    public DataTable LoadItemIds(IUserSession session, long objID, List<int> relIDs, List<int> childObjIDs, int lv, bool mode)
    {
        ICompositionLoadService _loadService = session.GetCustomService(typeof(ICompositionLoadService)) as ICompositionLoadService;
        List<ColumnDescriptor> col = new List<ColumnDescriptor>
        {
            new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_ID, AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.Index, SortOrders.NONE, 0),
        };
        DataTable reslt = null;
        session.EditingContextID = objID;
        DataTable dt = _loadService.LoadComplexCompositions(
        session.SessionGUID,
        new List<ObjInfoItem> { new ObjInfoItem(objID) },
        relIDs,
        childObjIDs,
        col,
        mode, // Режим получения данных (true - состав, false - применяемость)
        false, // Флаг группировки данных
        null, // Правило подбора версий
        null, // Условия на объекты при получении состава / применяемости
        //SystemGUIDs.filtrationLatestVersions, // Настройки фильтрации объектов (последние)
        SystemGUIDs.filtrationBaseVersions, // Настройки фильтрации объектов (базовые)
        //SystemGUIDs.filtrationLatestVersions,
        null,
        lv // Количество уровней для обработки (-1 = загрузка полного состава / применяемости)
        );
        if (dt == null)
            return reslt;
        else
            return dt;
    }
}
