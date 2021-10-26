using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;
using Misoft.V8Enterprise.Interfaces;

namespace MyIPSClient
{
    public class Script
    {
        public ICSharpScriptContext ScriptContext { get; set; }

        public void Execute(IActivity activity)
        {
            if (Debugger.IsAttached)
                Debugger.Break();

            ExportMode exportMode = ExportMode.ExportNotice;
            ExportSettings exportSettings = new ExportSettings();

            exportSettings.TypeObjectUseAsSectionOfSpecification = true;
            exportSettings.UseAssembleUnitFromTechprocess = true;
            exportSettings.UseLifeCycleAggred = false;
            exportSettings.UseLifeCycleDesign = false;
            exportSettings.UseLifeCyclePilotProduction = false;
            exportSettings.UseLifeCycleProduction = true;
            exportSettings.UseLifeCycleRemoved = false;
            exportSettings.UseLifeCycleStorage = false;
            exportSettings.Rule = RuleVariant.LatestVersionWithLifeCycle;
            exportSettings.CreateSpecForComplexMaterial = true;
            exportSettings.UsePositionInCompletingUnit = true;
            exportSettings.DontUseIsTechMapsInProductionLifeCycle = false;

            exportSettings.TMHSource = "IPS_DMZ";
            exportSettings.TMHVariant = "1";
            exportSettings.TMHPlantCode = "130";
            exportSettings.TMHExport = false;

            foreach (var attachment in activity.Attachments)
            {
                long idObjectIPS = attachment.ObjectID; //идентификатор версии извещения

                IV8Enterprise v8servise = activity.Session.GetCustomService(typeof(IV8Enterprise)) as IV8Enterprise;
                if (v8servise == null)
                {
                    throw new NotificationException("Сервис IV8Enterprise не найден");
                }

                IDBObject dBObject = activity.Session.GetObject(idObjectIPS, false);
                if (dBObject == null)
                {
                    return;
                }

                v8servise.ClearLog();
                byte[] result = v8servise.GetXMLOfIntermechObject(activity.Session.SessionGUID, idObjectIPS, exportMode, null, exportSettings);
                string xmlString = "";
                v8servise.ClearLog();

                using (var stream = CompressionHelper.Decompress(result))
                using (StreamReader reader = new StreamReader(stream, System.Text.Encoding.UTF8))
                {
                    xmlString = reader.ReadToEnd();
                }

                Misoft.V8Enterprise.ExchangeProviders.ComService1CProvider comProvider = new Misoft.V8Enterprise.ExchangeProviders.ComService1CProvider();

                comProvider.Settings.Mode = V8Mode.ServerMode;
                comProvider.Settings.ClasterName = "sql10-srv";
                comProvider.Settings.DatabaseName = "UKBMZerp_test22";
                comProvider.Settings.Username = "Мисофт";
                comProvider.Settings.Password = " ";
                //comProvider.Settings.PathToBase = Client.exportSettings.ComPath;
                comProvider.Settings.Version = V8Version.V83;
                comProvider.Settings.UseWindowsAuthorization = false;

                if (!comProvider.Initialize()) throw new System.Exception(comProvider.GetLastError());

                if (!comProvider.Connect()) throw new ExecutionEngineException(comProvider.GetLastError());


                Misoft.V8Enterprise.ExchangeProviders.ComObject1C DataProcessors = comProvider.DataProcessors["ипсОбменДанными"];
                if (comProvider.IsError) throw new System.Exception(comProvider.GetLastError());

                Misoft.V8Enterprise.ExchangeProviders.ComObject1C DataProcessorsObject = new Misoft.V8Enterprise.ExchangeProviders.ComObject1C(comProvider, DataProcessors.Run("Create"));
                if (comProvider.IsError) throw new System.Exception(comProvider.GetLastError());

                DataProcessorsObject["СтрокаXML"] = xmlString;
                if (comProvider.IsError) throw new System.Exception(comProvider.GetLastError());

                //выполнить загрузку
                object ResultOfExecDataProcessors = DataProcessorsObject.Run("Import");
                if (comProvider.IsError) throw new System.Exception(comProvider.GetLastError());

                DataProcessorsObject = null;
                DataProcessors = null;
                //comProvider.Disconnect();
            }



        }


    }



}
