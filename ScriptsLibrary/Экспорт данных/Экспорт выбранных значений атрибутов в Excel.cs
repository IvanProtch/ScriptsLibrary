using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;
using Intermech.Kernel.Search;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Microsoft.Office.Interop.Excel;

//Необходим скрипт, который будет записывать информацию по объектам
//из выборки в Excel файл.Каждый атрибут необходимо записывать 
//в отдельный столбец Excel файла. Необходимые атрибуты:
//«Идентификатор версии объекта», «Заголовок объекта», «Обозначение»,
//«Наименование», «Код АМТО». Тип объекта постоянно будет разный.
//По возможности не завязываться на определенное количество
//атрибутов т.к. количество может меняться. 

public class Script
{
    public ICSharpScriptContext ScriptContext { get; private set; }
    public string Error = string.Empty;

    private void SendToExcel(System.Data.DataTable table)
    {
        Application excelDoc = new Application();
        Workbook excelWb;
        Worksheet excelWs;

        excelDoc.Visible = true;
        excelDoc.Caption = table.TableName;

        excelWb = excelDoc.Workbooks.Add();
        excelWs = (Worksheet)excelWb.ActiveSheet;
        //excelWs.Name = ;

        try
        {
            //заполнение заголовка столбцов
            for (int i = 0; i < table.Columns.Count; i++)
            {
                excelWs.Cells[1, i + 1] = table.Columns[i].Caption;
                ((Range)excelWs.Cells[1, i + 1]).BorderAround2(Weight: XlBorderWeight.xlMedium);
            }


            //заполнение ячеек
            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    (excelWs.Cells[i + 2, j + 1]) = table.Rows[i].ItemArray[j].ToString();
                    ((Range)excelWs.Cells[i + 2, j + 1]).BorderAround2(Weight: XlBorderWeight.xlThin);
                }
            }

            //выравнивание
            for (int i = 0; i < table.Columns.Count; i++)
            {
                ((Range)excelWs.Columns[i + 1]).AutoFit();
            }

        }
        catch (Exception ex)
        {
            Error += (ex.Message) + Environment.NewLine;
        }
        finally
        {
            table.Dispose();
        }
    }

    private System.Data.DataTable SelectObjects(List<long> attrSelPossableValList, IUserSession UserSession)
    {
        IDBObjectCollection objects = UserSession.GetObjectCollection(-1);
        //Необходимые атрибуты: «Идентификатор версии объекта», «Заголовок объекта», «Обозначение», «Наименование», «Код АМТО».
        ColumnDescriptor[] columns =
        {
            new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_ID,
            AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.Name, SortOrders.NONE, 0),

            new ColumnDescriptor(ObligatoryObjectAttributes.CAPTION,
            AttributeSourceTypes.Object, ColumnContents.String, ColumnNameMapping.Name, SortOrders.ASC, 0),

            new ColumnDescriptor("Обозначение",
            AttributeSourceTypes.Object, ColumnContents.String, ColumnNameMapping.Name, SortOrders.NONE, 0),

            new ColumnDescriptor("Наименование",
            AttributeSourceTypes.Object, ColumnContents.String, ColumnNameMapping.Name, SortOrders.NONE, 0),

            new ColumnDescriptor("Код АМТО",
            AttributeSourceTypes.Object, ColumnContents.String, ColumnNameMapping.Name, SortOrders.NONE, 0),

        };
        ConditionStructure[] conds =
        {
            new ConditionStructure((int)ObligatoryObjectAttributes.F_OBJECT_ID, RelationalOperators.In, attrSelPossableValList.ToArray(),
            LogicalOperators.NONE, 0, true)
        };
        DBRecordSetParams dbrsp = new DBRecordSetParams(conds, columns);
        System.Data.DataTable tableObjs = new System.Data.DataTable();
        try
        {
            tableObjs = objects.Select(dbrsp);

        }
        catch (Exception ex)
        {
            Error += ex.Message + Environment.NewLine;
        }

        return tableObjs;
    }


    public void Execute(IActivity activity)
    {
        List<long> attachedItems = new List<long>();
        foreach (IAttachment item in activity.Attachments)
        {
            attachedItems.Add(item.ObjectID);
        }

        System.Data.DataTable dataTable = SelectObjects(attachedItems, activity.Session);
        SendToExcel(dataTable);

        if (Error.Length > 0)
            throw new NotificationException(Error);
    }

}

