using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;
using Intermech.Kernel.Search;
using System.Data;
using System.Diagnostics;
using Misoft.V8Enterprise.Catalogs;
using Misoft.V8Enterprise.ExchangeProviders;
using Misoft.V8Enterprise.Interfaces;

namespace MyIPSClient
{
    public class Script
    {
        const long selectionID = 45998523;
        const string PathToExport = @"D:\Mydoc\misoft\Рабочий стол\Формат ТМХ\";

        public ICSharpScriptContext ScriptContext { get; set; }

        public void Execute(IActivity activity)
        {
            
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

            //long idObjectIPS = 98038135; //идентификатор версии извещения

            try
            {

                foreach (var attachent in activity.Attachments)
                {

                    long idObjectIPS = attachent.ObjectID;

                    Package package = new Package();
                    package.StartExportTime = DateTime.Now;
                    package.TypeExchangeFuction = exportMode.ToString();
                    package.PluginVersion = "";

                    CatalogNotice notice = new CatalogNotice(activity.Session, idObjectIPS, null);
                    package.Add(notice);
                    package.ExportTime = DateTime.Now;
                    byte[] productXML = package.ToXmlArray();
                    string xmlString = "";

                    using (var stream = CompressionHelper.Decompress(productXML))
                    using (StreamReader reader = new StreamReader(stream, System.Text.Encoding.UTF8))
                    {
                        xmlString = reader.ReadToEnd();
                    }

                    ComService1CProvider comProvider = new ComService1CProvider();

                    comProvider.Settings.Mode = V8Mode.ServerMode;
                    comProvider.Settings.ClasterName = "sql10-srv";
                    comProvider.Settings.DatabaseName = "UKBMZerp_test22";
                    comProvider.Settings.Username = "Мисофт";
                    comProvider.Settings.Password = "123";
                    //comProvider.Settings.PathToBase = Client.exportSettings.ComPath;
                    comProvider.Settings.Version = V8Version.V83;
                    comProvider.Settings.UseWindowsAuthorization = false;
                    if (!comProvider.Initialize()) throw new System.Exception(comProvider.GetLastError());

                    ComObject1C DataProcessors = comProvider.DataProcessors["ипсОбменДанными"];
                    if (comProvider.IsError) throw new System.Exception(comProvider.GetLastError());

                    ComObject1C DataProcessorsObject = new ComObject1C(comProvider, DataProcessors.Run("Create"));
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
            catch (Exception ex)
            {
                throw new NotificationException(ex.Message);
            }



        }

    }



}
