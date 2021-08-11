//Проверка подписей «Директор производства», «Технический директор» у всех конструкторских ИИ во вложении процесса и связанных с ними.
//После выполнения выводит диалог с выбором о необходимости продолжения процесса или перехода по обратной связи
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
using System.Windows.Forms;

public class Script
{
    public ICSharpScriptContext ScriptContext { get; private set; }

    /// <summary>
    /// Наименование логической переменной процесса: если значение истинно, скрипт будет запущен
    /// </summary>
    private const string isExecutebleVar = "БМЗ_Выход шаг2";

    #region Константы

    private const int ConstrIIType = 2760;
    private const int signType = 1025;
    private const int criptSignType = 1101;

    #endregion Константы

    #region Типы связей
    private const int relTypeChangingByII = 1007;//Изменяется по извещению
    private const int relTypeSign = 1013; //Состав подписей
    #endregion Типы связей


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


    private string ActualSignsCheck(IUserSession UserSession, long obj, string signGraphValue)
    {
        string message = string.Empty;
        IDBObject o = UserSession.GetObject(obj);
        List<long> signs = LoadItemIds(UserSession, obj, new List<int> { relTypeChangingByII, relTypeSign },
        new List<int> { signType, criptSignType }, 1);

        int actualSignCount = 0;
        foreach (long sign in signs)
        {
            IDBObject signObj = UserSession.GetObject(sign);

            DateTime modifyDate = o.GetAttributeByGuid(new Guid("cad0013a-306c-11d8-b4e9-00304f19f545"
                /*Дата модификации содержимого объекта*/)).AsDateTime;

            DateTime signDate = signObj.GetAttributeByGuid(new Guid("cad014cb-306c-11d8-b4e9-00304f19f545"
                /*Дата подписания*/)).AsDateTime;

            string signGraph = signObj.GetAttributeByGuid(new Guid("cad00141-306c-11d8-b4e9-00304f19f545")).Description;

            //подписанные signGraphValue 
            if (modifyDate < signDate && signGraph == signGraphValue)
                actualSignCount++;
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
        else if(actualSignCount == 0)
            message += string.Format("\r\nДля {0} требуется обновить подпись в графе {1}.\r\n",
                o.NameInMessages, signGraphValue);

        return message;
    }

    public void Execute(IActivity activity)
    {
        if (Debugger.IsAttached)
            Debugger.Break();

        if ((bool)activity.Variables.Find(isExecutebleVar).TypedValue == false)
            return;

        IUserSession UserSession = activity.Session;
        string FinalMessage = string.Empty;

        foreach (IAttachment attachment in activity.Attachments)
        {
            if(MetaDataHelper.IsObjectTypeChildOf(attachment.ObjectType, ConstrIIType))
            {
                IDBEditingContextsService contextService = activity.Session.GetCustomService(typeof(IDBEditingContextsService)) as IDBEditingContextsService;
                EditingContextsObjectContainer contextContainer = contextService.GetEditingContextsObject(activity.Session.SessionGUID, attachment.ObjectID, true, true);

                var constIIs = contextContainer.GetContextsID()
                    .Where(ii => activity.Session.GetObject(ii).TypeID == ConstrIIType);

                foreach (var ii in constIIs)
                {
                    IDBObject II = UserSession.GetObject(ii);

                    #region Проверка актуальности подписей у извещений
                    string msg = string.Empty;

                    msg = ActualSignsCheck(UserSession, II.ObjectID, "Директор производства");
                    if (msg.Length > 0)
                        FinalMessage += msg;

                    msg = ActualSignsCheck(UserSession, II.ObjectID, "Технический директор");
                    if (msg.Length > 0)
                        FinalMessage += msg;

                    #endregion

                }
            }
        }
        if (FinalMessage.Length > 0)
        {
            FinalMessage += string.Format("\nПродолжить процесс?\n");
            DialogResult dlgRes = MessageBox.Show(FinalMessage, "Ошибка проверки подписей", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if(dlgRes == DialogResult.No)
                throw new NotificationException(string.Format("Процесс {0} вернулся по обратной связи.", activity.Process.Caption));
        }
    }
}