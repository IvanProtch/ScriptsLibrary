//Проверка ИИ исполнителя действия
//Выбор ИИ по исполнителю действия, проверка организации-источника, редактируемости объектов, актуальных подписей
using System;
using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;
using Intermech.Workflow;
using System.Xml.Linq;
using System.Xml;
using System.Diagnostics;
using Intermech.Interfaces.Contexts;
using System.Linq;
using System.Collections.Generic;
using Intermech.Interfaces.Compositions;
using Intermech.Kernel.Search;
using System.Data;
using Intermech;

public class Script0
{
    public ICSharpScriptContext ScriptContext { get; private set; }

    #region Константы

    #region ID Организаций-источников
    const long idBMZ = 3251710; // id объекта организации "БМЗ" 
    const long idKSK = 63822858; // id объекта организации ОП "КСК-Брянск"
    #endregion

    /// <summary>
    /// Технологический объект
    /// </summary>
    private const int techObjType = 1012;

    /// <summary>
    /// Техпроцесс базовый
    /// </summary>
    private const int baseTPType = 1117;

    /// <summary>
    /// Материалы
    /// </summary>
    private const int materialObjType = 1021;

    /// <summary>
    /// Тип извещения корневой
    /// </summary>
    private const int IIType = 1066;

    /// <summary>
    /// Тип извещения конструкторское
    /// </summary>
    private const int ConstrIIType = 2760;

    private const int TechIIType = 2761;
    private const int TechSvIIType = 2762;
    private const int signType = 1025;
    private const int criptSignType = 1101;

    private List<int> techObjTypeListAll;
    private List<int> materialObjTypeListAll;
    private List<int> TPListAll;
    private List<int> IIListAll;
    private List<int> TechIITypes;



    #endregion Константы

    #region Типы связей
    private const int relTypeChangingByII = 1007;//Изменяется по извещению
    private const int relTypeTechConsist = 1002; //Технологический состав
    private const int relTypeSign = 1013; //Состав подписей
    #endregion Типы связей

    /// <summary>
    /// Найти объект, который есть в контексте, но не найден в составе объекта II
    /// </summary>
    /// <param name="UserSession"></param>
    /// <param name="II">Родительский объект (извещение об изменении)</param>
    /// <param name="consistance">Его состав</param>
    /// <returns></returns>

    private List<long> NotConsistedInII_ObjIDs(IUserSession UserSession, long II, List<long> consistance)
    {
        // сервис для работы с контекстами
        IDBEditingContextsService contextService = UserSession.GetCustomService(typeof(IDBEditingContextsService)) as IDBEditingContextsService;

        EditingContextsObjectContainer contextContainer = contextService.GetEditingContextsObject(UserSession.SessionGUID, II, true, true);

        consistance = consistance
        .Where(element => UserSession.GetObject(element).ModificationID == contextContainer.ModificationID)
        .ToList();

        List<long> context = contextContainer.GetVersionsID(false, UserSession.UserID);

        List<long> result = context
        .Where(element => !consistance.Contains(element))
        .ToList();

        return result;
    }


    /// <summary>
    /// Получение списка ID состава/применяемости
    /// </summary>
    /// <param name="session"></param>
    /// <param name="objID"></param>
    /// <param name="relIDs"></param>
    /// <param name="childObjIDs"></param>
    /// <param name="lv">Количество уровней для обработки (-1 = загрузка полного состава / применяемости)</param>
    /// <returns></returns>

    private List<long> LoadItemIds(IUserSession session, long objID, List<int> relIDs, List<int> childObjIDs, int lv)
    {
        ICompositionLoadService _loadService = session.GetCustomService(typeof(ICompositionLoadService)) as ICompositionLoadService;
        List<ColumnDescriptor> col = new List<ColumnDescriptor>
        {
            new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_ID, AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.Index, SortOrders.NONE, 0),
        };
        List<long> reslt = new List<long>();
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
            return reslt;
        else
            return dt.Rows.OfType<DataRow>()
        .Select(element => (long)element[0]).ToList<long>();
    }


    private bool? FromOrg(IUserSession UserSession, long orgID, long elem)
    {
        IDBObject obj = UserSession.GetObject(elem);
        IDBAttribute orgIstAtr = obj.GetAttributeByName("Организация-источник");
        if (orgIstAtr.IsNull)
        {
            return null;
        }
        long org = (long)orgIstAtr.Value;
        if (org == orgID)
            return true;
        else
            return false;
    }

    private string ActualSignsCheck(IUserSession UserSession, long obj, string signGraphValue)
    {
        string message = string.Empty;
        IDBObject o = UserSession.GetObject(obj);
        List<long> signs = LoadItemIds(UserSession, obj, new List<int> { relTypeChangingByII, relTypeSign },
        new List<int> { signType, criptSignType }, 1);

        foreach (long sign in signs)
        {
            IDBObject signObj = UserSession.GetObject(sign);

            DateTime modifyDate = o.GetAttributeByGuid(new Guid("cad0013a-306c-11d8-b4e9-00304f19f545"
                /*Дата модификации содержимого объекта*/)).AsDateTime;

            DateTime signDate = signObj.GetAttributeByGuid(new Guid("cad014cb-306c-11d8-b4e9-00304f19f545"
                /*Дата подписания*/)).AsDateTime;

            string signGraph = signObj.GetAttributeByGuid(new Guid("cad00141-306c-11d8-b4e9-00304f19f545")).Description;
            string user = signObj.GetAttributeByGuid(new Guid("cad00143-306c-11d8-b4e9-00304f19f545")).Description;

            if (modifyDate > signDate && signGraph == signGraphValue)
            {
                message += string.Format("\r\nДля {0} требуется обновить подпись пользователя {1} в графе {2}.\r\n",
                    o.NameInMessages, user, signGraph);
            }
        }
        // Проверка наличия подписи в графе signGraph
        if (signs
            .Select(s => UserSession.GetObject(s))
            .Where(s => s.GetAttributeByGuid(new Guid("cad00141-306c-11d8-b4e9-00304f19f545")).Description == signGraphValue)
            .ToList().Count == 0)
        {
            message += string.Format("\r\nДля объекта {0} не задана ни одна подпись в графе {1}\r\n",
                o.NameInMessages, signGraphValue);
        }
        return message;

    }

    public void Execute(IActivity activity)
    {
        if (Debugger.IsAttached)
            Debugger.Break();

        IUserSession UserSession = activity.Session;
        string FinalMessage = string.Empty;
        // Инициализация списков типов объектов, используемых в запросах

        techObjTypeListAll = MetaDataHelper.GetObjectTypeChildrenIDRecursive(techObjType);
        materialObjTypeListAll = MetaDataHelper.GetObjectTypeChildrenIDRecursive(materialObjType);
        IIListAll = MetaDataHelper.GetObjectTypeChildrenIDRecursive(IIType);
        TPListAll = MetaDataHelper.GetObjectTypeChildrenIDRecursive(baseTPType);

        // Набор всех типов
        List<int> allObjectTypes = new List<int>();
        allObjectTypes.AddRange(techObjTypeListAll);
        allObjectTypes.AddRange(materialObjTypeListAll);
        allObjectTypes.AddRange(IIListAll);
        allObjectTypes.AddRange(TPListAll);

        TechIITypes = new List<int>();
        TechIITypes.Add(TechIIType);
        TechIITypes.Add(TechSvIIType);

        var recepient = activity.GetAttributeByGuid(new Guid("cad002ca-306c-11d8-b4e9-00304f19f545" /*Получатель*/));

        foreach (IAttachment attachment in activity.Attachments)
        {
            //Проходим только по конструкторским или технологическим ИИ
            if (MetaDataHelper.IsObjectTypeChildOf(attachment.ObjectType, ConstrIIType)
                || MetaDataHelper.IsObjectTypeChildOf(attachment.ObjectType, TechIIType)
            || MetaDataHelper.IsObjectTypeChildOf(attachment.ObjectType, TechSvIIType))
            {
                IDBEditingContextsService contextService = activity.Session.GetCustomService(typeof(IDBEditingContextsService)) as IDBEditingContextsService;
                EditingContextsObjectContainer contextContainer = contextService.GetEditingContextsObject(activity.Session.SessionGUID, attachment.ObjectID, true, true);

                List<long> techIIs = contextContainer.GetContextsID()
                    .Where(ii => TechIITypes.Contains(activity.Session.GetObject(ii).TypeID))
                    .Where(e => activity.Session.GetObject(e).CreatorID == recepient.AsInteger)
                    .ToList();

                List<List<long>> techIIsAllConsistance = new List<List<long>>();

                foreach (long techII in techIIs)
                {
                    //Получаем состав техн ИИ и заносим в список составов
                    List<long> consistance = LoadItemIds(UserSession, techII, new List<int> { relTypeChangingByII, relTypeTechConsist }, allObjectTypes, -1)
                        .Distinct().ToList<long>();
                    techIIsAllConsistance.Add(consistance);

                    #region Проверка актуальности подписей у извещений
                    string msg = string.Empty;
                    IDBObject IIobj = UserSession.GetObject(techII);

                    if (IIobj.ObjectType == TechIIType)
                    {
                        msg = ActualSignsCheck(UserSession, IIobj.ObjectID, "Составил (технология)");
                        if (msg.Length > 0)
                            FinalMessage += msg;
                    }

                    if (IIobj.ObjectType == TechSvIIType)
                    {
                        msg = ActualSignsCheck(UserSession, IIobj.ObjectID, "Составил (сварка)");
                        if(msg.Length > 0)
                            FinalMessage += msg;
                    }

                    #endregion
                }

                //Выполняем основные проверки ИИ
                //Проходим по всем технологическим ИИ
                for (int i = 0; i < techIIs.Count; i++)
                {
                    //Прошло ли ИИ проверку?
                    bool isChecked = true;
                    IDBObject II = UserSession.GetObject(techIIs[i]);

                    //Пропускаем дальнейшую проверку для чужих ИИ
                    if (FromOrg(UserSession, idBMZ, techIIs[i]) == false || FromOrg(UserSession, idKSK, techIIs[i]) == false)
                        continue;

                    //Проверяем редактируется ли само ИИ
                    if (II.CheckoutBy > 0)
                    {
                        IDBObject redactor = UserSession.GetObject(II.CheckoutBy);

                        FinalMessage += string.Format("\r\n[{0}] взял объект [{1}]|[{2}] на редактирование.\r\n",
                        redactor.NameInMessages, II.NameInMessages, II.ObjectID);
                        isChecked = false;
                    }
                    
                    //Проверка принадлежности организации БМЗ
                    if (FromOrg(UserSession, idBMZ, techIIs[i]) == null)
                    {
                        FinalMessage += string.Format("\r\n[{0}]|[{1}] значение атрибута 'организация-источник' пустое," +
                        " проверьте принадлежность ИИ БМЗ и укажите значение атрибута и актуализируйте.\r\n",
                        II.NameInMessages, II.ObjectID);
                        isChecked = false;
                        continue;
                    }

                    if (FromOrg(UserSession, idKSK, techIIs[i]) == null)
                    {
                        FinalMessage += string.Format("\r\n[{0}]|[{1}] значение атрибута 'организация-источник' пустое," +
                        " проверьте принадлежность ИИ БМЗ и укажите значение атрибута и актуализируйте.\r\n",
                        II.NameInMessages, II.ObjectID);
                        isChecked = false;
                        continue;
                    }
                    //Состав ИИ
                    List<long> consistance = techIIsAllConsistance[i];

                    //Проходим по составу
                    for (int ii_ind = 0; ii_ind < consistance.Count; ii_ind++)
                    {
                        IDBObject obj = UserSession.GetObject(consistance[ii_ind]);

                        IDBRelation relationII = UserSession.GetRelation(II.ObjectID, obj.ID);

                        if (relationII != null)
                        {
                            IDBAttribute reasonToEnterObj = relationII.GetAttributeByName("Цель включения объекта");
                            IDBAttribute customIDLC = relationII.GetAttributeByGuid(new Guid("cad01483-306c-11d8-b4e9-00304f19f545"
                                /*Идентификатор шага ЖЦ, на который будет переведен объект*/));

                            if ((reasonToEnterObj.Description == "Создание" || reasonToEnterObj.Description == "Изменение")
                                && Convert.ToInt32(customIDLC.Value) != 1058/*Шаг 'Согласование'*/)
                            {
                                isChecked = false;
                                FinalMessage += string.Format("Нельзя перевести объект {0} из {1} на шаг 'Согласование и утверждение'\n",
                                    obj.NameInMessages, II.NameInMessages);
                            }
                        }

                        if (obj.CheckoutBy > 0)
                        {
                            IDBObject redactor = UserSession.GetObject(obj.CheckoutBy);

                            FinalMessage += string.Format("\r\n[{0}] взял объект [{1}]|[{3}] из [{2}]|[{4}] на редактирование.\r\n",
                            redactor.NameInMessages, obj.NameInMessages, II.NameInMessages, obj.ObjectID, II.ObjectID);

                            isChecked = false;
                        }

                        #region Проверка подписей у объектов состава ИИ
                        List<int> techIIConsTypes = new List<int>()
                        {
                            MetaDataHelper.GetObjectTypeID("cad001da-306c-11d8-b4e9-00304f19f545" /*Заготовка*/),
                            MetaDataHelper.GetObjectTypeID("cad001e5-306c-11d8-b4e9-00304f19f545" /*Расцеховочный маршрут*/),
                            MetaDataHelper.GetObjectTypeID("7f6993e7-4b37-4dd3-a671-c9e571b3fd93" /*Группа инструментов*/),
                            MetaDataHelper.GetObjectTypeID("b3ec04e4-9d56-4494-a57b-766d10cdfe27" /*Группа материалов*/),
                            MetaDataHelper.GetObjectTypeID("cad00188-306c-11d8-b4e9-00304f19f545" /*Техпроцесс типовой*/),
                            MetaDataHelper.GetObjectTypeID("cad00187-306c-11d8-b4e9-00304f19f545" /*Техпроцесс единичный*/),
                            MetaDataHelper.GetObjectTypeID("cad00186-306c-11d8-b4e9-00304f19f545" /*Техпроцесс групповой*/)
                        };

                        List<int> techSvIIConsTypes = new List<int>()
                        {
                            MetaDataHelper.GetObjectTypeID("cad00188-306c-11d8-b4e9-00304f19f545" /*Техпроцесс типовой*/),
                            MetaDataHelper.GetObjectTypeID("cad00187-306c-11d8-b4e9-00304f19f545" /*Техпроцесс единичный*/),
                            MetaDataHelper.GetObjectTypeID("cad00186-306c-11d8-b4e9-00304f19f545" /*Техпроцесс групповой*/)
                        };
                        string msg = string.Empty;
                        if (II.ObjectType == TechIIType && techIIConsTypes.Contains(obj.ObjectType))
                        {
                            msg = ActualSignsCheck(UserSession, obj.ObjectID, "Разработал (технология)");
                            if (msg.Length > 0)
                                FinalMessage += msg;
                        }

                        if (II.ObjectType == TechSvIIType && techSvIIConsTypes.Contains(obj.ObjectType))
                        {
                            msg = ActualSignsCheck(UserSession, obj.ObjectID, "Разработал (сварка)");
                            if (msg.Length > 0)
                                FinalMessage += msg;
                        }
                        #endregion
                    }

                    //Получаем список объектов не найденных в составе, но присутствующих в контексте ИИ
                    List<long> nConsistedObjIds = NotConsistedInII_ObjIDs(UserSession, techIIs[i], techIIsAllConsistance[i]);
                    foreach (long nConsistedObj in nConsistedObjIds)
                    {
                        IDBObject obj = UserSession.GetObject(nConsistedObj);

                        FinalMessage += string.Format("\r\nОбъект [{0}]|[{2}] присутствует в контексте редактирования [{1}]|[{3}]," +
                        " но не обнаружен в составе.\r\n",
                        obj.NameInMessages, II.NameInMessages, obj.ObjectID, II.ObjectID);

                        if (obj.CheckoutBy > 0)
                        {
                            IDBObject redactor = UserSession.GetObject(obj.CheckoutBy);

                            FinalMessage += string.Format("[{0}] взял объект [{1}]|[{2}] на редактирование.\r\n",
                            redactor.NameInMessages, obj.NameInMessages, obj.ObjectID);
                        }
                        isChecked = false;
                    }

                }
            }
        }
        if (FinalMessage.Length > 0)
        {
            throw new NotificationException(FinalMessage);
        }
    }
}