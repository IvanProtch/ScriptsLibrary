//      Скрипт проверки наличия аутентичных файлов

//Во вложении БП идет ИИ конструкторское. в составе конструкторского ИИ наобходимо
//найти типы объектов: "Спецификация", "Чертеж детали Inventor", "Сборочный чертеж Inventor".
//В объектах проверить наличие аутентичных файлов (тип файла) и их актуальность.
//Для проверки актуальности необходимо сравнить дату изменения аутентичного файла
//и дату модификации объекта. Аутентичный файл считается актуальным если дата
//модификации объекта меньше или равна дате изменения аутентичного файла.
//Так же наличие и актуальность аутентичного файла необходимо проверить и на конструкторском ИИ.    

using System;
using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;
using Intermech.Workflow;
using System.Xml.Linq;
using System.Xml;
using System.Diagnostics;
using System.IO;
using Intermech;
using System.Collections.Generic;
using Intermech.Interfaces.Compositions;
using Intermech.Kernel.Search;
using System.Data;
using System.Linq;

public class Script
{

    public ICSharpScriptContext ScriptContext { get; private set; }

    public void Execute(IActivity activity)
    {
        if (Debugger.IsAttached)
            Debugger.Break();

        string error = string.Empty;

        List<int> relations = new List<int>()
        {
            1, 1007
        };
        List<int> types = new List<int>() 
        { 
            MetaDataHelper.GetObjectTypeID(new Guid("cad00902-306c-11d8-b4e9-00304f19f545" /*Сборочные чертежи Inventor*/)),
            MetaDataHelper.GetObjectTypeID(new Guid("cad00133-306c-11d8-b4e9-00304f19f545" /*Спецификации*/)),
            MetaDataHelper.GetObjectTypeID(new Guid("cad00909-306c-11d8-b4e9-00304f19f545" /*Чертежи деталей Inventor*/))
        };

        foreach (var attachment in activity.Attachments)
        {
            if(attachment.ObjectType == 2760 /*Конструкторское извещение*/)
            {
                var ob = attachment.Object;
                var childsAndSelf = LoadItemChils(activity.Session, attachment.ObjectID, relations, types, -1);
                childsAndSelf.Add(attachment.Object);
                foreach (var item in childsAndSelf)
                {
                    var fileAttr = item.GetAttributeByGuid(new Guid(SystemGUIDs.attributeFile));
                    bool hasAuthentical = false;
                    for (int i = 0; i < fileAttr.ValuesCount; i++)
                    {
                        MemoryStream ms = new MemoryStream();
                        BlobProcReader bpr = new BlobProcReader(item.ObjectID, AttributableElements.Object, MetaDataHelper.GetAttributeTypeID(new Guid(SystemGUIDs.attributeFile)), i, 0, ms, null, null);
                        bpr.ReadData();
                        BlobInformation file = bpr.BlobInformation;

                        if (file.FileType == FileTypes.ftAuthentical)
                        {
                            hasAuthentical = true;
                            if (item.ModifyDate > file.ModifyDate)
                                error += string.Format("У объекта {0} обнаружен неактуальный аутентичный файл {1}.\n\r", item.NameInMessages, file.FileName);
                        }
                    }
                    if(!hasAuthentical)
                        error += string.Format("У объекта {0} не найден аутентичный файл.\n\r", item.NameInMessages);
                }
            }
        }

        if (error.Length > 0)
            throw new ClientException(error);


    }

    private List<IDBObject> LoadItemChils(IUserSession session, long objID, List<int> relIDs, List<int> childObjIDs, int lv)
    {
        ICompositionLoadService _loadService = session.GetCustomService(typeof(ICompositionLoadService)) as ICompositionLoadService;
        List<ColumnDescriptor> col = new List<ColumnDescriptor>
        {
            new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_ID, AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.Index, SortOrders.NONE, 0),
        };
        session.EditingContextID = objID;
        DataTable dt = _loadService.LoadComplexCompositions(
        session.SessionGUID,
        new List<ObjInfoItem> { new ObjInfoItem(objID) },
        relIDs,
        childObjIDs,
        col,
        true, // Режим получения данных (true - состав, false - применяемость)
        false, // Флаг группировки данных
        null, // Правило подбора версий
        null, // Условия на объекты при получении состава / применяемости
              //SystemGUIDs.filtrationLatestVersions, // Настройки фильтрации объектов (последние)
              //SystemGUIDs.filtrationBaseVersions, // Настройки фильтрации объектов (базовые)
        Intermech.SystemGUIDs.filtrationLatestVersions,
        null,
        lv // Количество уровней для обработки (-1 = загрузка полного состава / применяемости)
        );
        if (dt == null)
            return null;
        else
            return dt.Rows.OfType<DataRow>()
        .Select(element => session.GetObject((long)element[0]))
        .ToList();
    }
}