using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using _38_39Conversion.ExcelObjects;
using System.Xml;
using _38_39Conversion.Utils;
using System.Windows.Forms;

namespace _38_39Conversion.XmlGenerationFiles
{
    class _920Xml
    {
        public static void build920Dm(_920Module _920, string filePath)
        {
            XmlDocument doc = new XmlDocument();
            XmlDeclaration xmldecl = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
            XmlNode dmodule = doc.CreateElement("dmodule");

            XmlNode identAndStatusSection = doc.CreateElement("identAndStatusSection");

            XmlNode dmAddress = doc.CreateElement("dmAddress");

            XmlNode dmIdent = doc.CreateElement("dmIdent");

            //populate dmCode--------------------------
            XmlNode dmCode = XmlUtils.BuildDmRef(_920._920DMC, doc);

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
            techName.InnerText = _920._920DmcTitle.Substring(0, _920._920DmcTitle.IndexOf(" - "));

            XmlNode infoName = doc.CreateElement("infoName");
            infoName.InnerText = _920._920DmcTitle.Substring(_920._920DmcTitle.IndexOf(" - ")).TrimStart(new char[] { ' ', '-' });
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
            XmlNode refs = bulid920Refs(_920, doc);

            content.AppendChild(refs);

            XmlNode procedure = build920Procedure(_920, doc);

            content.AppendChild(procedure);

            dmodule.AppendChild(content);

            doc.AppendChild(dmodule);

            XmlElement root = doc.DocumentElement;
            doc.InsertBefore(xmldecl, root);
            doc.Save(filePath + "/" + _920._920DMC + ".xml");
        }

        public static XmlNode bulid920Refs(_920Module  _920, XmlDocument doc)
        {
            XmlNode refs = doc.CreateElement("refs");

            XmlNode dmRef1 = doc.CreateElement("dmRef");

            XmlNode dmRefIdent1 = doc.CreateElement("dmRefIdent");

            string _520dmc = _920._920DMC.Replace("920","520");

            XmlNode dmCode1 = XmlUtils.BuildDmRef(_520dmc,doc);

            dmRefIdent1.AppendChild(dmCode1);

            dmRef1.AppendChild(dmRefIdent1);

            XmlNode dmRefAddressItems1 = doc.CreateElement("dmRefAddressItems");

            XmlNode dmTitle1 = doc.CreateElement("dmTitle");

            XmlNode techName1 = doc.CreateElement("techName");
            techName1.InnerText = _920._920DmcTitle.Substring(0, _920._920DmcTitle.IndexOf(" - "));

            XmlNode infoName1 = doc.CreateElement("infoName");
            infoName1.InnerText = "Remove procedure";

            dmTitle1.AppendChild(techName1);
            dmTitle1.AppendChild(infoName1);

            dmRefAddressItems1.AppendChild(dmTitle1);

            dmRef1.AppendChild(dmRefAddressItems1);

            //sep
            XmlNode dmRef2 = doc.CreateElement("dmRef");

            XmlNode dmRefIdent2 = doc.CreateElement("dmRefIdent");

            string _720dmc = _920._920DmcTitle.Replace("920", "720");

            XmlNode dmCode2 = XmlUtils.BuildDmRef(_520dmc, doc);

            dmRefIdent2.AppendChild(dmCode2);

            dmRef2.AppendChild(dmRefIdent2);

            XmlNode dmRefAddressItems2 = doc.CreateElement("dmRefAddressItems");

            XmlNode dmTitle2 = doc.CreateElement("dmTitle");
            XmlNode techName2 = doc.CreateElement("techName");
            techName2.InnerText = _920._920DmcTitle.Substring(0, _920._920DmcTitle.IndexOf(" - "));

            XmlNode infoName2 = doc.CreateElement("infoName");
            infoName2.InnerText = "Install procedure";

            dmTitle2.AppendChild(techName2);
            dmTitle2.AppendChild(infoName2);

            dmRefAddressItems2.AppendChild(dmTitle2);

            dmRef2.AppendChild(dmRefAddressItems2);

            refs.AppendChild(dmRef1);
            refs.AppendChild(dmRef2);

            return refs;
        }

        public static XmlNode build920Procedure(_920Module _920, XmlDocument doc)
        {
            XmlNode procedure = doc.CreateElement("procedure");

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

            procedure.AppendChild(preliminaryRqmts);

            XmlNode mainProcedure = doc.CreateElement("mainProcedure");

            XmlNode proceduralStep520 = build920ProceduralStep(_920, doc, "520", "PS0001","Remove Procedure");
            XmlNode proceduralStep720 = build920ProceduralStep(_920, doc, "720", "PS0002", "Install Procedure");
            mainProcedure.AppendChild(proceduralStep520);
            mainProcedure.AppendChild(proceduralStep720);

            procedure.AppendChild(mainProcedure);

            XmlNode closeRqmts = doc.CreateElement("closeRqmts");

            XmlNode reqCondGroupCls = doc.CreateElement("reqCondGroup");
            XmlNode noCondsCls = doc.CreateElement("noConds");
            reqCondGroupCls.AppendChild(noCondsCls);

            closeRqmts.AppendChild(reqCondGroupCls);

            procedure.AppendChild(closeRqmts);

            return procedure;
        }

        public static XmlNode build920ProceduralStep(_920Module _920, XmlDocument doc, string infoCode, string idValue, string infoName)
        {
            XmlNode proceduralStep = doc.CreateElement("proceduralStep");
            XmlAttribute id = doc.CreateAttribute("id");
            id.InnerText = idValue;
            proceduralStep.Attributes.Append(id);

            XmlNode para = doc.CreateElement("para");

            XmlNode dmRef1 = doc.CreateElement("dmRef");

            XmlNode dmRefIdent1 = doc.CreateElement("dmRefIdent");

            string _520dmc = _920._920DMC.Replace("920", infoCode);

            XmlNode dmCode1 = XmlUtils.BuildDmRef(_520dmc, doc);

            dmRefIdent1.AppendChild(dmCode1);

            dmRef1.AppendChild(dmRefIdent1);

            XmlNode dmRefAddressItems1 = doc.CreateElement("dmRefAddressItems");

            XmlNode dmTitle1 = doc.CreateElement("dmTitle");

            XmlNode techName1 = doc.CreateElement("techName");
            techName1.InnerText = _920._920DmcTitle.Substring(0, _920._920DmcTitle.IndexOf(" - "));

            XmlNode infoName1 = doc.CreateElement("infoName");
            infoName1.InnerText = infoName;

            dmTitle1.AppendChild(techName1);
            dmTitle1.AppendChild(infoName1);

            dmRefAddressItems1.AppendChild(dmTitle1);

            dmRef1.AppendChild(dmRefAddressItems1);

            para.AppendChild(dmRef1);

            proceduralStep.AppendChild(para);

            return proceduralStep;
        }
    }
}
