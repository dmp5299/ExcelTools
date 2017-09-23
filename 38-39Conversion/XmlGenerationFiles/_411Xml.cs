using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using _38_39Conversion.ExcelObjects;
using _38_39Conversion.Utils;
using System.Windows.Forms;
using System.Threading;
using System.ComponentModel;

namespace _38_39Conversion.XmlGenerationFiles
{
    public class _411Xml
    {
        public static void build411Dms(List<_411Module> _411ModuleData, BackgroundWorker worker)
        {
            for (int i=0;i<_411ModuleData.Count;i++)
            {
                _920Xml.build920Dm(_411ModuleData[i]._920Element, _411ModuleData[i].excelPath);
                build411Dm(_411ModuleData[i]);
                worker.ReportProgress(i+1);
                Thread.Sleep(100);
            }
            MessageBox.Show("done");
        }

        public static void build411Dm(_411Module _411)
        {
            XmlDocument doc = new XmlDocument();
            XmlDeclaration xmldecl = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlNode dmodule = doc.CreateElement("dmodule");

            XmlNode identAndStatusSection = doc.CreateElement("identAndStatusSection");

            XmlNode dmAddress = doc.CreateElement("dmAddress");

            XmlNode dmIdent = doc.CreateElement("dmIdent");

            //populate dmCode--------------------------
            XmlNode dmCode = XmlUtils.BuildDmRef(_411._411DMC, doc);

            dmIdent.AppendChild(dmCode);

            //Language element------------------------------------------------

            XmlNode language = doc.CreateElement("language");

            XmlAttribute countryIsoCode = doc.CreateAttribute("countryIsoCode");
            countryIsoCode.InnerText = "US";
            language.Attributes.Append(countryIsoCode);

            XmlAttribute languageIsoCode = doc.CreateAttribute("languageIsoCode");
            languageIsoCode.InnerText = "sx";
            language.Attributes.Append(languageIsoCode);

            dmIdent.AppendChild(language);

            //issueInfo element------------------------------------------------

            XmlNode issueInfo = doc.CreateElement("issueInfo");

            XmlAttribute inWork = doc.CreateAttribute("inWork");
            inWork.InnerText = "00";
            issueInfo.Attributes.Append(inWork);

            XmlAttribute issueNumber = doc.CreateAttribute("issueNumber");
            issueNumber.InnerText = "000";
            issueInfo.Attributes.Append(issueNumber);

            dmIdent.AppendChild(issueInfo);

            dmAddress.AppendChild(dmIdent);

            //End of dmIdent-----------------------------------------------------

            XmlNode dmAddressItems = doc.CreateElement("dmAddressItems");

            //IssueDate----------------------------------------------------------

            XmlNode issueDate = doc.CreateElement("issueDate");

            XmlAttribute day = doc.CreateAttribute("day");
            day.InnerText = "15";
            issueDate.Attributes.Append(day);

            XmlAttribute month = doc.CreateAttribute("month");
            month.InnerText = "05";
            issueDate.Attributes.Append(month);

            XmlAttribute year = doc.CreateAttribute("year");
            year.InnerText = "2015";
            issueDate.Attributes.Append(year);

            dmAddressItems.AppendChild(issueDate);

            //dmTitle-----------------------------------------------------------

            XmlNode dmTitle = doc.CreateElement("dmTitle");

            XmlNode techName = doc.CreateElement("techName");
            techName.InnerText = _411._411DmcTitle.Substring(0, _411._411DmcTitle.IndexOf(" - "));

            XmlNode infoName = doc.CreateElement("infoName");
            infoName.InnerText = _411._411DmcTitle.Substring(_411._411DmcTitle.IndexOf(" - ")).TrimStart(new char[] { ' ','-'});
            dmTitle.AppendChild(techName);
            dmTitle.AppendChild(infoName);

            dmAddressItems.AppendChild(dmTitle);

            dmAddress.AppendChild(dmAddressItems);

            identAndStatusSection.AppendChild(dmAddress);

            //dmStatus------------------------------------------------------

            XmlNode dmStatus = doc.CreateElement("dmStatus");

            XmlAttribute issueType = doc.CreateAttribute("issueType");
            issueType.InnerText = "new";
            dmStatus.Attributes.Append(issueType);

            XmlNode security = doc.CreateElement("security");

            XmlAttribute securityClassification = doc.CreateAttribute("securityClassification");
            securityClassification.InnerText = "01";
            security.Attributes.Append(securityClassification);

            dmStatus.AppendChild(security);

            XmlNode dataRestrictions = doc.CreateElement("dataRestrictions");
            XmlNode restrictionInstructions = doc.CreateElement("restrictionInstructions");
            XmlNode dataDistribution = doc.CreateElement("dataDistribution");

            restrictionInstructions.AppendChild(dataDistribution);

            XmlNode exportControl = doc.CreateElement("exportControl");
            XmlNode exportRegistrationStmt = doc.CreateElement("exportRegistrationStmt");
            XmlNode simplePara = doc.CreateElement("simplePara");
            exportRegistrationStmt.AppendChild(simplePara);
            exportControl.AppendChild(exportRegistrationStmt);
            restrictionInstructions.AppendChild(exportControl);
            dataRestrictions.AppendChild(restrictionInstructions);
            dmStatus.AppendChild(dataRestrictions);

            XmlNode responsiblePartnerCompany = doc.CreateElement("responsiblePartnerCompany");

            XmlAttribute enterpriseCode = doc.CreateAttribute("enterpriseCode");
            enterpriseCode.InnerText = "";
            responsiblePartnerCompany.Attributes.Append(enterpriseCode);
            dmStatus.AppendChild(responsiblePartnerCompany);

            XmlNode originator = doc.CreateElement("originator");

            XmlAttribute enterpriseCode1 = doc.CreateAttribute("enterpriseCode");
            enterpriseCode1.InnerText = "";
            originator.Attributes.Append(enterpriseCode1);
            dmStatus.AppendChild(originator);

            XmlNode applic = doc.CreateElement("applic");
            XmlNode displayText = doc.CreateElement("displayText");
            XmlNode simplePara1 = doc.CreateElement("simplePara");
            simplePara1.InnerText = "All";
            displayText.AppendChild(simplePara1);
            applic.AppendChild(displayText);
            dmStatus.AppendChild(applic);

            XmlNode brexDmRef = doc.CreateElement("brexDmRef");
            XmlNode dmRef = doc.CreateElement("dmRef");
            XmlAttribute xlink_actuate = doc.CreateAttribute("xlink:actuate");
            xlink_actuate.InnerText = "onRequest";
            dmRef.Attributes.Append(xlink_actuate);
            XmlAttribute xlink_href = doc.CreateAttribute("xlink:href");
            xlink_href.InnerText = "URN:S1000D:DMC-S1000DBIKE-AAA-D00-00-00-00AA-022A-D_005";
            dmRef.Attributes.Append(xlink_href);
            XmlAttribute xlink_show = doc.CreateAttribute("xlink:show");
            xlink_show.InnerText = "replace";
            dmRef.Attributes.Append(xlink_show);
            XmlAttribute xlink_type = doc.CreateAttribute("xlink:type");
            xlink_type.InnerText = "simple";
            dmRef.Attributes.Append(xlink_type);

            XmlNode dmRefIdent = doc.CreateElement("dmRefIdent");

            XmlNode dmCode1 = doc.CreateElement("dmCode");

            XmlAttribute assyCode1 = doc.CreateAttribute("assyCode");
            assyCode1.InnerText = "00";
            dmCode1.Attributes.Append(assyCode1);

            XmlAttribute disassyCode1 = doc.CreateAttribute("disassyCode");
            assyCode1.InnerText = "00";
            dmCode1.Attributes.Append(disassyCode1);

            XmlAttribute disassyCodeVariant1 = doc.CreateAttribute("disassyCodeVariant");
            disassyCodeVariant1.InnerText = "AA";
            dmCode1.Attributes.Append(disassyCodeVariant1);

            XmlAttribute infoCode1 = doc.CreateAttribute("infoCode");
            infoCode1.InnerText = "022";
            dmCode1.Attributes.Append(infoCode1);

            XmlAttribute infoCodeVariant1 = doc.CreateAttribute("infoCodeVariant");
            infoCodeVariant1.InnerText = "A";
            dmCode1.Attributes.Append(infoCodeVariant1);

            XmlAttribute itemLocationCode1 = doc.CreateAttribute("itemLocationCode");
            itemLocationCode1.InnerText = "D";
            dmCode1.Attributes.Append(itemLocationCode1);

            XmlAttribute modelIdentCode1 = doc.CreateAttribute("modelIdentCode");
            modelIdentCode1.InnerText = "S1000DBIKE";
            dmCode1.Attributes.Append(modelIdentCode1);

            XmlAttribute subSubSystemCode1 = doc.CreateAttribute("subSubSystemCode");
            subSubSystemCode1.InnerText = "0";
            dmCode1.Attributes.Append(subSubSystemCode1);

            XmlAttribute subSystemCode1 = doc.CreateAttribute("subSystemCode");
            subSystemCode1.InnerText = "0";
            dmCode1.Attributes.Append(subSystemCode1);

            XmlAttribute systemCode1 = doc.CreateAttribute("systemCode");
            systemCode1.InnerText = "D00";
            dmCode1.Attributes.Append(systemCode1);

            XmlAttribute systemDiffCode1 = doc.CreateAttribute("systemDiffCode");
            systemDiffCode1.InnerText = "AAA";
            dmCode1.Attributes.Append(systemDiffCode1);

            dmRefIdent.AppendChild(dmCode1);

            XmlNode issueInfo1 = doc.CreateElement("issueInfo");
            XmlAttribute inWork1 = doc.CreateAttribute("inWork");
            inWork1.InnerText = "00";
            issueInfo1.Attributes.Append(inWork1);

            XmlAttribute issueNumber1 = doc.CreateAttribute("issueNumber");
            issueNumber1.InnerText = "005";
            issueInfo1.Attributes.Append(issueNumber1);
            dmRefIdent.AppendChild(issueInfo1);

            dmRef.AppendChild(dmRefIdent);
            brexDmRef.AppendChild(dmRef);
            dmStatus.AppendChild(brexDmRef);

            XmlNode qualityAssurance = doc.CreateElement("qualityAssurance");
            XmlNode unverified = doc.CreateElement("unverified");
            qualityAssurance.AppendChild(unverified);
            dmStatus.AppendChild(qualityAssurance);


            identAndStatusSection.AppendChild(dmStatus);

            dmodule.AppendChild(identAndStatusSection);

            XmlNode content = doc.CreateElement("content");

            XmlNode faultIsolation = buildFaultIsolationProcedures(_411.FaultIsolationElements, doc);

            content.AppendChild(faultIsolation);

            dmodule.AppendChild(content);

            doc.AppendChild(dmodule);

            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmldecl, root);
            doc.Save(_411.excelPath + "/" + _411._411DMC + ".xml");
        }

        public static XmlNode buildFaultIsolationProcedures(List<FaultIsolation> faultIsolationProcedures, XmlDocument doc)
        {
            XmlNode faultIsolation = doc.CreateElement("faultIsolation");

            foreach (FaultIsolation f in faultIsolationProcedures)
            {
                XmlNode faultIsolationProcedure = doc.CreateElement("faultIsolationProcedure");
                XmlAttribute id = doc.CreateAttribute("id");
                id.InnerText = f.FaultIsolationProcedureId;
                faultIsolationProcedure.Attributes.Append(id);

                XmlNode fault = doc.CreateElement("fault");

                XmlAttribute faultCode = doc.CreateAttribute("faultCode");
                faultCode.InnerText = f.FaultCode;

                fault.Attributes.Append(faultCode);

                faultIsolationProcedure.AppendChild(fault);

                XmlNode faultDescr = doc.CreateElement("faultDescr");
                XmlNode descr = doc.CreateElement("descr");
                descr.InnerText = f.MaintenanceTaskName;
                faultDescr.AppendChild(descr);

                faultIsolationProcedure.AppendChild(faultDescr);

                XmlNode isolationProcedure = doc.CreateElement("isolationProcedure");

                XmlNode preliminaryRqmts = doc.CreateElement("preliminaryRqmts");

                XmlNode reqCondGroup = doc.CreateElement("reqCondGroup");
                XmlNode noConds = doc.CreateElement("noConds");
                reqCondGroup.AppendChild(noConds);

                preliminaryRqmts.AppendChild(reqCondGroup);

                XmlNode reqSupportEquips = doc.CreateElement("reqSupportEquips");
                XmlNode noSupportEquips = doc.CreateElement("noSupportEquips");
                reqSupportEquips.AppendChild(noSupportEquips);

                preliminaryRqmts.AppendChild(reqSupportEquips);

                XmlNode reqSupplies = doc.CreateElement("reqSupplies");
                XmlNode noSupplies = doc.CreateElement("noSupplies");
                reqSupplies.AppendChild(noSupplies);

                preliminaryRqmts.AppendChild(reqSupplies);

                XmlNode reqSpares = doc.CreateElement("reqSpares");
                XmlNode noSpares = doc.CreateElement("noSpares");
                reqSpares.AppendChild(noSpares);

                preliminaryRqmts.AppendChild(reqSpares);

                XmlNode reqSafety = doc.CreateElement("reqSafety");
                XmlNode noSafety = doc.CreateElement("noSafety");
                reqSafety.AppendChild(noSafety);

                preliminaryRqmts.AppendChild(reqSafety);

                isolationProcedure.AppendChild(preliminaryRqmts);

                XmlNode isolationMainProcedure = doc.CreateElement("isolationMainProcedure");

                XmlNode isolationProcedureEnd = doc.CreateElement("isolationProcedureEnd");

                XmlAttribute isolationProcedureEndId = doc.CreateAttribute("id");
                isolationProcedureEndId.InnerText = f.FaultIsolationProcedureId.Replace("FI", "IE");
                isolationProcedureEnd.Attributes.Append(isolationProcedureEndId);

                XmlNode action = doc.CreateElement("action");

                XmlNode dmRef = doc.CreateElement("dmRef");

                XmlNode dmRefIdent = doc.CreateElement("dmRefIdent");

                XmlNode dmCode = XmlUtils.BuildDmRef(f._920DMC, doc);

                dmRefIdent.AppendChild(dmCode);

                dmRef.AppendChild(dmRefIdent);

                XmlNode dmRefAddressItems = doc.CreateElement("dmRefAddressItems");

                XmlNode dmTitle = doc.CreateElement("dmTitle");

                XmlNode techName = doc.CreateElement("techName");
                techName.InnerText = f._920DmcTitle.Substring(0, f._920DmcTitle.IndexOf(" - "));

                XmlNode infoName = doc.CreateElement("infoName");
                infoName.InnerText = f._920DmcTitle.Substring(f._920DmcTitle.IndexOf(" - ")).TrimStart(new char[] { ' ', '-' });

                dmTitle.AppendChild(techName);
                dmTitle.AppendChild(infoName);

                dmRef.AppendChild(dmRefAddressItems);

                action.AppendChild(dmRef);

                isolationProcedureEnd.AppendChild(action);

                isolationMainProcedure.AppendChild(isolationProcedureEnd);

                isolationProcedure.AppendChild(isolationMainProcedure);

                XmlNode closeRqmts = doc.CreateElement("closeRqmts");

                XmlNode reqCondGroupCls = doc.CreateElement("reqCondGroup");
                XmlNode noCondsCls = doc.CreateElement("noConds");
                reqCondGroupCls.AppendChild(noCondsCls);

                closeRqmts.AppendChild(reqCondGroupCls);

                isolationProcedure.AppendChild(closeRqmts);

                faultIsolationProcedure.AppendChild(isolationProcedure);

                faultIsolation.AppendChild(faultIsolationProcedure);
            }
            return faultIsolation;
        }
    }
}
