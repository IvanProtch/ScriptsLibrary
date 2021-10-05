//Актуализация связанных техн ИИ_адм версия

//При возникновении системной ошибки, сообщение не отправляется в виде письма, а выводится в окне
using Intermech;
using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;
using Intermech.Kernel.Search;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System;
using System.Xml;
using System.Diagnostics;
using Intermech.Workflow;
using Intermech.Interfaces.Compositions;

    public class Script
    {
        public ICSharpScriptContext ScriptContext { get; private set; }

        #region Константы
        /// <summary>
        /// Технологический объект
        /// </summary>
        const int techObjType = 1012;
        /// <summary>
        /// Техпроцесс базовый
        /// </summary>
        const int baseTPType = 1117;
        /// <summary>
        /// Материалы
        /// </summary>
        const int materialObjType = 1021;
        /// <summary>
        /// Тип извещения корневой
        /// </summary>
        const int IIType = 1066;
        /// <summary>
        /// Тип извещения конструкторское
        /// </summary>
        const int ConstrIIType = 2760;
        /// <summary>
        /// Взаимосвязанный контекст для ИИ
        /// </summary>
        const int connectedContectId = 12677;

        List<int> techObjTypeListAll;
        List<int> materialObjTypeListAll;
        List<int> TPListAll;
        List<int> IIListAll;
        List<int> TechIITypes;
        /// <summary>
        /// Объекты, которые не нужно переводить на согласование
        /// </summary>
        private List<int> objTypesToExcludeLCstepChanging = new List<int> { 1231/*оснастка*/, 1128/*материал*/, 1096/*материал составной*/, 1118/*оборудование*/, 1094/*модель оборудования*/ };
        const int TechIIType = 2761;
        const int TechSvIIType = 2762;

        const int lcStep_Sogl = 1023;
        const int lcStep_Final = 1024;
        const int lcStep_Development = 1021;


        private const int TechObj_lcStep_Sogl = 1056;
        private const int TechObj_lcStep_Final = 1058;
        private const int TechObj_lcStep_Development = 1050;
        #endregion

    #region Типы связей
    const int relTypeChangingByII = 1007;//Изменяется по извещению
        const int relTypeDocNaIzd = 1004; //Документация на изделие
        const int relTypeConsist = 1; //Состав изделия
        const int relTypeTechConsist = 1002; //Технологический состав
        const int relTypeTechData = 1011; //Собираемые технологические данные
        const int relTypeTechConectScetch = 1012; //Технологическая связь с эскизом
        const int relTypeTP = 1036; //Сквозной ТП
        const int relTypeMO = 1060; //МО обработки изделия
        const int relTypeConection = 0; //Простая связь
        #endregion

        #region ID Организаций-источников

        const long idBMZ = 3251710; // id объекта организации "БМЗ" 
        const long idKSK = 63822858; // id объекта организации ОП "КСК-Брянск"



        #endregion


        /// <summary>
        /// Производит выбор объектов из БД по значению атрибута-селектора
        /// </summary>
        /// <param name="mainType">Тип искомого объекта</param>
        /// <param name="attrSelector"></param>
        /// <param name="attrSelValue"></param>
        /// <returns></returns>

        private List<long> SelectObjects(IUserSession UserSession, int mainType,
        int attrSelector, long attrSelValue)
        {
            IDBObjectCollection objects = UserSession.GetObjectCollection(mainType);

            ColumnDescriptor[] columns =
            {
            new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_ID,
            AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.ID, SortOrders.ASC, 0),
            new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_TYPE,
            AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.ID, SortOrders.NONE, 0)
        };
            ConditionStructure[] conds =
            {
            new ConditionStructure((int)attrSelector, RelationalOperators.Equal, (int)attrSelValue, LogicalOperators.NONE, 0, true)
        };
            DBRecordSetParams dbrsp = new DBRecordSetParams(conds, columns, 0, null, QueryConsts.All);

            DataTable tableObjs = objects.Select(dbrsp);
            List<long> idsL = tableObjs.Rows.
            OfType<DataRow>().
            Select(a => (long)a[0]).
            ToList();

            return idsL;
        }

        /// <summary>
        /// Получает связанные технологические ИИ с выбранным ИИ любого типа
        /// </summary>
        /// <param name="UserSession"></param>
        /// <param name="constrIIid"></param>
        /// <returns></returns>

        private List<long> GetConnectedTechIIs(IUserSession UserSession, long IIid)
        {
            IDBObject mainII = UserSession.GetObject(IIid);
            IDBAttribute connectedContextAtr = mainII.GetAttributeByID(connectedContectId);

            List<long> techIIs = SelectObjects(UserSession, IIType,
            connectedContectId, Convert.ToInt64(connectedContextAtr.Value));

            techIIs = techIIs
            .Where(ii => TechIITypes.Contains(UserSession.GetObject(ii).TypeID))
            //.Where(ii => ii != IIid)
            .ToList();

            return techIIs;
        }

        private bool? FromOrg(long orgID, IUserSession UserSession, long elem)
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


        public void Execute(IActivity activity)
        {
        if (Debugger.IsAttached)
            Debugger.Break();
            IUserSession UserSession = activity.Session;

            // Инициализация списков типов объектов, используемых в запросах
            IIListAll = MetaDataHelper.GetObjectTypeChildrenIDRecursive(IIType);
            // Набор всех типов
            TechIITypes = new List<int>();
            TechIITypes.Add(TechIIType);
            TechIITypes.Add(TechSvIIType);

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

            string FinalMessage = string.Empty;

            foreach (IAttachment attachment in activity.Attachments)
            {
                //Проходим только по конструкторским или технологическим ИИ
                if (MetaDataHelper.IsObjectTypeChildOf(attachment.ObjectType, ConstrIIType)
                    || MetaDataHelper.IsObjectTypeChildOf(attachment.ObjectType, TechIIType)
                || MetaDataHelper.IsObjectTypeChildOf(attachment.ObjectType, TechSvIIType))
                {
                    //Получаем список, связанных с главным ИИ, техн ИИ
                    List<long> techIIs = GetConnectedTechIIs(UserSession, attachment.ObjectID);

                    List <List<long>> techIIsAllConsistance = new List<List<long>>();

                    //Получаем состав техн ИИ и заносим в список составов
                    foreach (long techII in techIIs)
                    {
                        List<long> consistance = LoadItemIds(UserSession, techII,
                        new List<int> { relTypeChangingByII, relTypeTechConsist },
                        allObjectTypes, -1).Distinct().ToList<long>();

                        techIIsAllConsistance.Add(consistance);
                    }

                    if (techIIs.Count > 0)
                    {
                        //Проходим по всем технологическим ИИ
                        for (int i = 0; i < techIIs.Count; i++)
                        {
                            bool isChecked = true;

                            IDBObject II = UserSession.GetObject(techIIs[i]);

                            //Проверка принадлежности организации
                            if ((FromOrg(idBMZ, UserSession, techIIs[i]) == null || FromOrg(idKSK, UserSession, techIIs[i]) == null))
                            {
                                FinalMessage += string.Format("\r\n[{0}]|[{1}] значение атрибута 'организация-источник' пустое," +
                                " проверьте принадлежность ИИ БМЗ (или ОП \"КСК - Брянск\"), укажите значение атрибута и актуализируйте.\r\n\r\n",
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
                                        && Convert.ToInt32(customIDLC.Value) != 1058)
                                    {
                                        isChecked = false;
                                        FinalMessage += string.Format("Нельзя перевести объект {0} из {1} на шаг 'Актуализация'\n",
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
                            }

                            //Переводим ИИ на шаг "актуализация"
                            if (II.LCStep == lcStep_Sogl)
                            {
                                if (isChecked == true
                                && (FromOrg(idBMZ, UserSession, techIIs[i]) == true || FromOrg(idKSK, UserSession, techIIs[i]) == true))
                                {
                                    try
                                    {
                                        II.LCStep = lcStep_Final;
                                    }
                                    catch (KernelException exc)
                                    {
                                        FinalMessage += string.Format("При переводе {0} на шаг 'Актуализация' возникла системная ошибка: \r\n" +
                                        "Источник: {1} \r\n" +
                                        "Сообщение: {2} \r\n\r\n",
                                        II.NameInMessages, exc.Source, exc.Message);
                                    }
                                }
                            }
                            else if ((II.LCStep != lcStep_Sogl) && ((II.LCStep != lcStep_Final))
                                && (FromOrg(idBMZ, UserSession, techIIs[i]) == true || FromOrg(idKSK, UserSession, techIIs[i]) == true))
                            {
                                FinalMessage += string.Format("\r\nДля перевода на шаг 'Актуализация' [{0}]|[{1}] необходимо," +
                                " чтобы ИИ находилось на шаге 'Согласование и утверждение'.\r\n",
                                II.NameInMessages, II.ObjectID);
                                isChecked = false;
                            }

                            //Переводим объекты состава на шаг "актуализация"
                            foreach (long item in consistance)
                            {
                                IDBObject obj = UserSession.GetObject(item);
                                if (!objTypesToExcludeLCstepChanging.Contains(obj.TypeID))
                                {
                                    if (obj.LCStep == TechObj_lcStep_Sogl)
                                    {
                                        try
                                        {
                                            obj.LCStep = TechObj_lcStep_Final;
                                        }
                                        catch (KernelException exc)
                                        {
                                            FinalMessage += string.Format("При переводе {0}({3}) на шаг 'Актуализация' возникла системная ошибка: \r\n" +
                                            "Источник: {1}; \r\n" +
                                            "Сообщение: {2}; \r\n",
                                            obj.NameInMessages, exc.Source, exc.Message, item);
                                        }
                                    }
                                }
                            }   
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