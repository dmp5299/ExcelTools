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
                worker.ReportProgress(i + 1);
                //Thread.Sleep(5);
                
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
                languageIsoCode.InnerText = "en";
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
                if(_411._411DmcTitle.IndexOf(" - ")  > -1)
                {
                    techName.InnerText = _411._411DmcTitle.Substring(0, _411._411DmcTitle.IndexOf(" - "));
                }
                else if(_411._411DmcTitle.IndexOf("-") > -1)
                {
                    techName.InnerText = _411._411DmcTitle.Substring(0, _411._411DmcTitle.IndexOf("- "));
                }
                else
                {
                    throw new Exception("Error in formatting of dcm title: " + _411._411DmcTitle);
                }

                XmlNode infoName = doc.CreateElement("infoName");
                if (_411._411DmcTitle.IndexOf(" - ") > -1)
                {
                    infoName.InnerText = _411._411DmcTitle.Substring(_411._411DmcTitle.IndexOf(" - ")).TrimStart(new char[] { ' ', '-' });
                }
                else if (_411._411DmcTitle.IndexOf("-") > -1)
                {
                    infoName.InnerText = _411._411DmcTitle.Substring(_411._411DmcTitle.IndexOf("- ")).TrimStart(new char[] { ' ', '-' });
                }
                else
                {
                    throw new Exception("Error in formatting of dcm title: " + _411._411DmcTitle);
                }
                
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

            XmlNode applicCrossRefTableRef = doc.CreateElement("applicCrossRefTableRef");
                XmlNode applicCrossRefTableRefDmRef = doc.CreateElement("dmRef");
                XmlNode applicCrossRefTableRefDmRefDmRefIdent = doc.CreateElement("dmRefIdent");
                XmlNode applicCrossRefTableRefDmCode = XmlUtils.BuildDmRef("HH60W-A-00-00-0000-00AAA-00WA-A", doc);

            applicCrossRefTableRefDmRefDmRefIdent.AppendChild(applicCrossRefTableRefDmCode);
            applicCrossRefTableRefDmRef.AppendChild(applicCrossRefTableRefDmRefDmRefIdent);

            XmlNode dmRefAddressItems = doc.CreateElement("dmRefAddressItems");

            XmlNode applicCrossRefTableRefDmTitle = doc.CreateElement("dmTitle");

            XmlNode applicCrossRefTableRefTechName = doc.CreateElement("techName");
            applicCrossRefTableRefTechName.InnerText = "Combat Rescue Helicopter (CRH)";
            applicCrossRefTableRefDmTitle.AppendChild(applicCrossRefTableRefTechName);

            XmlNode applicCrossRefTableRefInfoName = doc.CreateElement("infoName");
            applicCrossRefTableRefInfoName.InnerText = "Applicability Cross-reference Table (ACT)";
            applicCrossRefTableRefDmTitle.AppendChild(applicCrossRefTableRefInfoName);

            dmRefAddressItems.AppendChild(applicCrossRefTableRefDmTitle);
            applicCrossRefTableRefDmRef.AppendChild(dmRefAddressItems);

            applicCrossRefTableRef.AppendChild(applicCrossRefTableRefDmRef);

            dmStatus.AppendChild(applicCrossRefTableRef);

            XmlNode responsiblePartnerCompany = doc.CreateElement("responsiblePartnerCompany");

                XmlAttribute enterpriseCode = doc.CreateAttribute("enterpriseCode");
                enterpriseCode.InnerText = "78286";
                responsiblePartnerCompany.Attributes.Append(enterpriseCode);
                dmStatus.AppendChild(responsiblePartnerCompany);
                XmlNode originator = doc.CreateElement("originator");

                XmlAttribute enterpriseCode1 = doc.CreateAttribute("enterpriseCode");
                enterpriseCode1.InnerText = "78286";
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

                XmlNode dmRefIdent = doc.CreateElement("dmRefIdent");

            XmlNode dmCode1 = XmlUtils.BuildDmRef("HH60W-A-00-00-0000-0000A-022A-D",doc);

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

            XmlNode brexDmRefAddressItems = doc.CreateElement("dmRefAddressItems");

            XmlNode brexDmRefDmtitle = doc.CreateElement("dmTitle");

            XmlNode brexDmRefTechName = doc.CreateElement("techName");

            brexDmRefTechName.InnerText = "Combat Rescue Helicopter (CRH";

            brexDmRefDmtitle.AppendChild(brexDmRefTechName);

            XmlNode brexDmRefInfoName = doc.CreateElement("infoName");

            brexDmRefInfoName.InnerText = "Business rule exchange (BREX)";

            brexDmRefDmtitle.AppendChild(brexDmRefInfoName);

            brexDmRefAddressItems.AppendChild(brexDmRefDmtitle);

            dmRef.AppendChild(brexDmRefAddressItems);

            brexDmRef.AppendChild(dmRef);
                dmStatus.AppendChild(brexDmRef);

                XmlNode qualityAssurance = doc.CreateElement("qualityAssurance");
                XmlNode unverified = doc.CreateElement("unverified");
                qualityAssurance.AppendChild(unverified);
                dmStatus.AppendChild(qualityAssurance);


                identAndStatusSection.AppendChild(dmStatus);

                dmodule.AppendChild(identAndStatusSection);

                XmlNode content = doc.CreateElement("content");

            XmlNode refs = doc.CreateElement("refs");

            XmlNode contentDmRef = doc.CreateElement("dmRef");

            XmlNode contentDmRefIdent = doc.CreateElement("dmRefIdent");

            XmlNode contentDmcode = XmlUtils.BuildDmRef(_411.FaultIsolationElements[0]._920DMC,doc);

            contentDmRefIdent.AppendChild(contentDmcode);

            contentDmRef.AppendChild(contentDmRefIdent);

            XmlNode contentDmRefAddressItems = doc.CreateElement("dmRefAddressItems");

            XmlNode contentDmtitle = doc.CreateElement("dmTitle");
            
            XmlNode contenTechName = doc.CreateElement("techName");

            contenTechName.InnerText = _411.FaultIsolationElements[0]._920DmcTitle.Substring(0, _411.FaultIsolationElements[0]._920DmcTitle.IndexOf(" - "));

            contentDmtitle.AppendChild(contenTechName);

            XmlNode contentInfo = doc.CreateElement("infoName");

            contentInfo.InnerText = _411.FaultIsolationElements[0]._920DmcTitle.Substring(_411.FaultIsolationElements[0]._920DmcTitle.IndexOf(" - ")).TrimStart(new char[] { ' ', '-' });

            contentDmtitle.AppendChild(contentInfo);

            contentDmRefAddressItems.AppendChild(contentDmtitle);

            contentDmRef.AppendChild(contentDmRefAddressItems);

            refs.AppendChild(contentDmRef);

            content.AppendChild(refs);

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
            XmlNode faultReporting = doc.CreateElement("faultReporting");

            foreach (FaultIsolation f in faultIsolationProcedures)
            {
                XmlNode isolatedFault = doc.CreateElement("isolatedFault");

                XmlAttribute id = doc.CreateAttribute("id");
                id.InnerText = f.FaultIsolationProcedureId;
                isolatedFault.Attributes.Append(id);

                XmlAttribute faultCode = doc.CreateAttribute("faultCode");
                faultCode.InnerText = f.FaultCode;
                isolatedFault.Attributes.Append(faultCode);

                XmlNode faultDescr = doc.CreateElement("faultDescr");

                XmlNode descr = doc.CreateElement("descr");

                descr.InnerText = f.FailureName;

                faultDescr.AppendChild(descr);

                isolatedFault.AppendChild(faultDescr);

                XmlNode locateAndRepair = doc.CreateElement("locateAndRepair");

                XmlNode locateAndRepairLruItem = doc.CreateElement("locateAndRepairLruItem");

                XmlNode lru = doc.CreateElement("lru");

                XmlNode name = doc.CreateElement("name");

                name.InnerText = f.Name;

                lru.AppendChild(name);

                locateAndRepairLruItem.AppendChild(lru);

                XmlNode repair = doc.CreateElement("repair");

                XmlNode refs = doc.CreateElement("refs");

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

                dmRefAddressItems.AppendChild(dmTitle);

                dmRef.AppendChild(dmRefAddressItems);

                refs.AppendChild(dmRef);

                repair.AppendChild(refs);

                locateAndRepairLruItem.AppendChild(repair);

                locateAndRepair.AppendChild(locateAndRepairLruItem);

                isolatedFault.AppendChild(locateAndRepair);

                faultReporting.AppendChild(isolatedFault);
            }
            return faultReporting;
        }
    }
}
