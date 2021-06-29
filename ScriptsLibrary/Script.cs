using System;
using Intermech.Interfaces;
using Intermech.Expert.Scenarios;
using Intermech.Interfaces.Document;
using Intermech.Project;
using Intermech.Project.Client;
using System.Windows.Forms;
using System.Text;

public class Script : ICustomReportScenario
{

    private ImDocumentData doc = null;

    private void Write(string id, string text)
    {
        TextData td = doc.FindNode(id) as TextData;
        if (td != null)
            td.AssignText(text, false, false, false);
    }

    private void Write(DocumentTreeNode parent, string tplid, string text)
    {
        TextData td = parent.FindFirstNodeFromTemplate_Recursive(tplid) as TextData;
        if (td != null)
            td.AssignText(text, false, false, false);
    }

    private DocumentTreeNode FindNode(string name, DocumentTreeNode doc = null)
    {
        if (doc == null)
            doc = this.doc;

        DocumentTreeNode node = doc.FindNode(name);
        if (node == null)
            throw new Exception(string.Format("Node \"{0}\" not found!", name));
        return node;
    }

    private string GetIndentString(Task t)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append(' ', t.IndentLevel);
        return sb.ToString();
    }

    private void SetIndent(DocumentTreeNode parent, string tplid, float indent)
    {
        TextData td = parent.FindFirstNodeFromTemplate_Recursive(tplid) as TextData;
        if (td != null)
        {
            ParagraphFormat pf = td.ParagraphFormat.Clone();
            pf.IdentLeft = indent;
            td.SetParagraphFormat(pf, false, false);
        }
    }

    public bool Execute(IUserSession session, ImDocumentData document, Int64[] objectIDs)
    {
        if (objectIDs.Length > 0)
        {
            //подключимся к активному проекту, если имеется
            ClientProject project = ProjectEditorForm.CurrentProject;

            //активного проекта нет, загрузим по поданному идентификатору
            if (project == null)
            {
                project = new ClientProject();
                project.Load(objectIDs[0], false);
            }

            doc = document;

            doc.Name = string.Format("Отчет по проекту \"{0}\"", project.Name);
            Write("title", project.Name);
            if (project.Filter != null)
                Write("date", "Задачи по фильтру \"" + project.Filter.Name + "\" на " + DateTime.Now.ToString("g"));
            else
                Write("date", "Все задачи на " + DateTime.Now.ToString("g"));

            DocumentTreeNode table = FindNode("table");
            DocumentTreeNode row = FindNode("row", doc.Template);

            int completedtasks = 0;
            int needcompletion = 0;
            foreach (Task t in project.Tasks)
            {
                if (t.IsProjectSummaryTask)
                    continue;

                //задачи, отфильтрованные фильтром в отчет не пишем
                if (t.HiddenByFilter)
                    continue;

                DocumentTreeNode node = row.CloneFromTemplate(true, true);
                table.AddChildNode(node, false, false);

                //Write(node, "name", GetIndentString(t) + t.Name);
                Write(node, "name", t.Name);
                SetIndent(node, "name", .1f + .3f * t.IndentLevel);

                Write(node, "start", project.FormatDateTime(t.Start));
                Write(node, "finish", project.FormatDateTime(t.Finish));

                double pp = t.PlannedPercentCompleted;
                Write(node, "plannedpercent", string.Format("{0:0.#}%", pp));
                if (pp >= 100)
                    needcompletion++;

                Write(node, "percent", string.Format("{0:0.#}%", t.PercentCompleted));
                if (t.PercentCompleted >= 100)
                    completedtasks++;

                Write(node, "users", string.Join(", ", t.Assignments.UserNames.ToArray()));
            }

            string s1 = "Наименование: " + project.Name;
            s1 += "\r\nПлановое начало: " + project.FormatDateTime(project.Start);
            string fs = "-";
            if (!project.FactStart.Equals(DateTime.MinValue))
                fs = project.FormatDateTime(project.FactStart);
            s1 += "\r\nФактическое начало: " + fs;
            s1 += "\r\nПлановое завершение: " + project.FormatDateTime(project.Finish);
            s1 += "\r\nРуководитель: " + project.ChiefName;

            string s2 = "Фактический прогресс: " + string.Format("{0:0.#}%", project.PercentCompleted);
            s2 += "\r\nПлановый прогресс: " + string.Format("{0:0.#}%", project.PlannedPercentCompleted);
            s2 += "\r\nВсего задач в проекте: " + project.Tasks.Count.ToString();
            s2 += "\r\nДолжно быть выполнено: " + needcompletion.ToString();
            s2 += "\r\nФактически выполнено: " + completedtasks.ToString();

            Write("summary1", s1);
            Write("summary2", s2);

            doc.UpdateLayout(0, true, false);
        }
        return true;
    }
}