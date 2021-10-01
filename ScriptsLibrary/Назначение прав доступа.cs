using System;
using System.Xml;
using System.Data;
using System.Linq;
using System.Collections.Generic;

using Intermech;
using Intermech.Kernel;
using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;
using Intermech.Interfaces.Contexts;

public class Script
{
    public ICSharpScriptContext ScriptContext { get; private set; }

    public void Execute(IActivity activity)
    {
        if (System.Diagnostics.Debugger.IsAttached)
            System.Diagnostics.Debugger.Break();

        IUserSession session = activity.Session;
        List<long> userGroupIDs = new List<long>() { 3254492/*Все пользователи БМЗ*/, 62874086/*КСК-БМЗ ТБ2 (ОП "КСК-Брянск")*/};
        List<int> objtypeECO = ScriptContext.MetaDataHelper.GetObjectTypeChildrenIDRecursive(new Guid(Intermech.SystemGUIDs.objtypeECO /*Извещения*/));

        // сервис для работы с контекстами
        IDBEditingContextsService contextService = session.GetCustomService(typeof(IDBEditingContextsService)) as IDBEditingContextsService;

        int t = 0;

        long[] ecoIDs = activity.Attachments.Where(a => objtypeECO.Contains(a.ObjectType)).Select(a => a.ObjectID).ToArray();
        foreach (long ecoID in ecoIDs)
        {
            // получаю содержимое контекста
            EditingContextsObjectContainer contextContainer = contextService.GetEditingContextsObject(session.SessionGUID, ecoID, true, true);
            List<long> contextIds = contextContainer != null ? contextContainer.GetVersionsID(ecoID, session.UserID) : new List<long>();
            // добавляем ИИ для назначения прав
            contextIds.Add(ecoID);

            foreach (long objContxtID in contextIds)
            {
                IDBObject objContxt = session.GetObject(objContxtID);

                IDBSecurity sec = (IDBSecurity)objContxt;

                ActionProperties[] actPr = null;
                QuickObjectInfo[] qObjInfo = null;

                // получаем таблицу прав доступа
                DataTable dtAccList = sec.GetAccessList(out actPr, out qObjInfo);

                // если право уже есть - пропускаем
                bool rightsEnabled = userGroupIDs
                    .All(group =>
                    dtAccList.Rows.OfType<DataRow>()
                        .Any(dataRow =>
                        group == Convert.ToInt64(dataRow[Consts.F_USER_ID])
                        && (ActionType)dataRow[Consts.F_RIGHT_ID] == ActionType.View
                        && (AccessType)dataRow[Consts.F_RIGHT_TYPE] == AccessType.Grant));

                if (rightsEnabled)
                    continue;
                else
                {
                    dtAccList.Rows.Clear();

                    foreach (var userGroupID in userGroupIDs)
                    {
                        // добавляем новую строку
                        DataRow dr = dtAccList.NewRow();

                        dr[Consts.F_CATEGORY_ID] = objContxt.ObjectID;
                        dr[Consts.F_CATEGORY_TYPE] = Consts.CategoryObjectVersion;
                        dr[Consts.F_RIGHT_ID] = ActionType.View;
                        dr[Consts.F_USER_ID] = userGroupID; // ид. юзера, которому назначили права
                        dr[Consts.F_PARENT_KEY] = 0;
                        dr[Consts.F_KEY] = 0;
                        dr[Consts.F_RIGHT_TYPE] = AccessType.Grant;
                        dr[Consts.F_OWNER_ID] = session.UserID;
                        dr[Consts.F_END_DATE] = DBNull.Value;
                        dr[Consts.F_BEGIN_DATE] = DBNull.Value;

                        dtAccList.Rows.Add(dr);
                    }

                    dtAccList.AcceptChanges();

                    // сохраняем
                    sec.SetAccess(dtAccList);
                }
            }
        }
    }
}
