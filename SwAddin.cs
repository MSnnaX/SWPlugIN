using System;
using System.Runtime.InteropServices;
using System.Reflection;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;
using SolidWorksTools;
using SolidWorksTools.File;
using System.Collections.Generic;
using System.IO;

namespace MSnnaXSwAddIn
{
    [Guid("9c79970d-ac3a-4a4f-b5de-9452422f26ca"), ComVisible(true)]
    [SwAddinAttribute(
        Description = "MSnnaXSwAddIn Application",
        Title = "MSnnaXSwAddIn",
        LoadAtStartup = true
        )]
    //iSwApp.SendMsgToUser();
    //System.Windows.Forms.MessageBox.Show("");
    //iSwApp.SetSelectionFilter((int)swSelectType_e.swSelDIMENSIONS, true);
    public class SwAddin : ISwAddin
    {
        #region Local Variables
        ISldWorks iSwApp = null;
        ICommandManager iCmdMgr = null;
        int addinID = 0;
        BitmapHandler iBmpHad;

        public const int mainCmdGroupID = 5;
        public const int flyoutGroupID = 91;

        //command ID of changing the decimal of dimension
        public const int mainItemID00 = 0;
        public const int mainItemID01 = 1;
        public const int mainItemID02 = 2;
        public const int mainItemID03 = 3;
        //command ID of changing the number of dimension
        public const int mainItemID10 = 4;
        public const int mainItemID12 = 5;
        public const int mainItemID13 = 6;
        public const int mainItemID14 = 7;
        public const int mainItemID15 = 8;
        public const int mainItemID16 = 9;
        public const int mainItemID17 = 10;
        public const int mainItemID18 = 11;
        public const int mainItemID19 = 12;
        //command ID of changing the tolerance of dimension
        public const int mainItemID20 = 13;
        public const int mainItemID21 = 14;
        public const int mainItemID22 = 15;
        public const int mainItemID23 = 16;
        //command ID of changing the text(THR) of dimension
        public const int mainItemID30 = 17;
        //command ID of inserting the Surface Finish Symbol
        public const int mainItemID40 = 18;
        public const int mainItemID41 = 19;
        public const int mainItemID42 = 20;
        public const int mainItemID43 = 21;
        //command ID of exporting files to current directory
        public const int mainItemID50 = 22;
        public const int mainItemID51 = 23;
        public const int mainItemID52 = 24;
        public const int mainItemID53 = 25;

        //SolidWorks.Interop.sldworks.SldWorks SwEventPtr = null;

        // Public Properties
        public ISldWorks SwApp
        {
            get { return iSwApp; }
        }
        public ICommandManager CmdMgr
        {
            get { return iCmdMgr; }
        }

        #endregion

        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute
            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion

            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }

            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);

                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        #region ISwAddin Implementation
        public SwAddin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;
            
            //Setup callbacks
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            #region Setup the Command Manager
            iCmdMgr = iSwApp.GetCommandManager(cookie);
            AddCommandMgr();
            #endregion

            //SwEventPtr = (SolidWorks.Interop.sldworks.SldWorks)iSwApp;

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(iCmdMgr);
            iCmdMgr = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp);
            iSwApp = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }
        #endregion

        #region UI Methods
        public void AddCommandMgr()
        {
            ICommandGroup cmdGroup;
            string Title = "MSnnaXSwAddIn", ToolTip = "MSnnaXSwAddIn";
            if (iBmpHad == null)
                iBmpHad = new BitmapHandler();
            Assembly thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());
            
            //command Index of changing the decimal of dimension
            int cmdIndex00, cmdIndex01, cmdIndex02, cmdIndex03;
            int cmdIndex10, cmdIndex12, cmdIndex13, cmdIndex14, cmdIndex15, cmdIndex16, cmdIndex17, cmdIndex18, cmdIndex19;
            int cmdIndex20, cmdIndex21, cmdIndex22, cmdIndex23;
            int cmdIndex30;
            int cmdIndex40, cmdIndex41, cmdIndex42, cmdIndex43;
            int cmdIndex50, cmdIndex51, cmdIndex52, cmdIndex53;

            bool ignorePrevious = false;
            int cmdGroupErr = 0;

            object registryIDs;
            //get the ID information stored in the registry
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);
            int[] knownIDs = new int[26] { mainItemID00, mainItemID01, mainItemID02, mainItemID03,
                mainItemID10, mainItemID12, mainItemID13, mainItemID14, mainItemID15, mainItemID16, mainItemID17, mainItemID18, mainItemID19,
                mainItemID20, mainItemID21, mainItemID22, mainItemID23,
                mainItemID30,
                mainItemID40, mainItemID41, mainItemID42, mainItemID43,
                mainItemID50, mainItemID51, mainItemID52, mainItemID53
            };

            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs)) //if the IDs don't match, reset the commandGroup
                {
                    ignorePrevious = true;
                }
            }

            cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);
            cmdGroup.LargeIconList = iBmpHad.CreateFileFromResourceBitmap("SWPlugIN.ToolbarL.bmp", thisAssembly);
            cmdGroup.SmallIconList = iBmpHad.CreateFileFromResourceBitmap("SWPlugIN.ToolbarS.bmp", thisAssembly);
            cmdGroup.LargeMainIcon = iBmpHad.CreateFileFromResourceBitmap("SWPlugIN.MainIconL.bmp", thisAssembly);
            cmdGroup.SmallMainIcon = iBmpHad.CreateFileFromResourceBitmap("SWPlugIN.MainIconS.bmp", thisAssembly);

            //menuToolbarOption = 1:display in menu;2:display in toolbar;3:display in both;
            int mTbO = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            cmdIndex00 = cmdGroup.AddCommandItem2("Decimal0", -1, "", "Dem D0", 00, "Decimal0", "EnableInDrw", mainItemID00, mTbO);
            cmdIndex01 = cmdGroup.AddCommandItem2("Decimal1", -1, "", "Dem D1", 01, "Decimal1", "EnableInDrw", mainItemID01, mTbO);
            cmdIndex02 = cmdGroup.AddCommandItem2("Decimal2", -1, "", "Dem D2", 02, "Decimal2", "EnableInDrw", mainItemID02, mTbO);
            cmdIndex03 = cmdGroup.AddCommandItem2("Decimal3", -1, "", "Dem D3", 03, "Decimal3", "EnableInDrw", mainItemID03, mTbO);
            
            cmdIndex10 = cmdGroup.AddCommandItem2("DimNumb0", -1, "", "DN 0x", 04, "DimNumb0", "EnableInDrw", mainItemID10, mTbO);
            cmdIndex12 = cmdGroup.AddCommandItem2("DimNumb2", -1, "", "DN 2x", 05, "DimNumb2", "EnableInDrw", mainItemID12, mTbO);
            cmdIndex13 = cmdGroup.AddCommandItem2("DimNumb3", -1, "", "DN 3x", 06, "DimNumb3", "EnableInDrw", mainItemID13, mTbO);
            cmdIndex14 = cmdGroup.AddCommandItem2("DimNumb4", -1, "", "DN 4x", 07, "DimNumb4", "EnableInDrw", mainItemID14, mTbO);
            cmdIndex15 = cmdGroup.AddCommandItem2("DimNumb5", -1, "", "DN 5x", 08, "DimNumb5", "EnableInDrw", mainItemID15, mTbO);
            cmdIndex16 = cmdGroup.AddCommandItem2("DimNumb6", -1, "", "DN 6x", 09, "DimNumb6", "EnableInDrw", mainItemID16, mTbO);
            cmdIndex17 = cmdGroup.AddCommandItem2("DimNumb7", -1, "", "DN 7x", 10, "DimNumb7", "EnableInDrw", mainItemID17, mTbO);
            cmdIndex18 = cmdGroup.AddCommandItem2("DimNumb8", -1, "", "DN 8x", 11, "DimNumb8", "EnableInDrw", mainItemID18, mTbO);
            cmdIndex19 = cmdGroup.AddCommandItem2("DimNumb9", -1, "", "DN 9x", 12, "DimNumb9", "EnableInDrw", mainItemID19, mTbO);

            cmdIndex20 = cmdGroup.AddCommandItem2("ToleranceN0", -1, "", "Tol N0", 13, "ToleranceN0", "EnableInDrw", mainItemID20, mTbO);
            cmdIndex21 = cmdGroup.AddCommandItem2("ToleranceH7", -1, "", "Tol H7", 14, "ToleranceH7", "EnableInDrw", mainItemID21, mTbO);
            cmdIndex22 = cmdGroup.AddCommandItem2("Toleranceg6", -1, "", "Tol g6", 15, "Toleranceg6", "EnableInDrw", mainItemID22, mTbO);
            cmdIndex23 = cmdGroup.AddCommandItem2("Tolerance01", -1, "", "Tol 01", 16, "Tolerance01", "EnableInDrw", mainItemID23, mTbO);

            cmdIndex30 = cmdGroup.AddCommandItem2("ADDTHR", -1, "", "ADD THR", 17, "ADDTHR", "EnableInDrw", mainItemID30, mTbO);

            cmdIndex40 = cmdGroup.AddCommandItem2("SurFinSymN0", -1, "", "SurFinSym N0", 18, "SurFinSymN0", "EnableInDrw", mainItemID40, mTbO);
            cmdIndex41 = cmdGroup.AddCommandItem2("SurFinSym04", -1, "", "SurFinSym 04", 19, "SurFinSym04", "EnableInDrw", mainItemID41, mTbO);
            cmdIndex42 = cmdGroup.AddCommandItem2("SurFinSym08", -1, "", "SurFinSym 08", 20, "SurFinSym08", "EnableInDrw", mainItemID42, mTbO);
            cmdIndex43 = cmdGroup.AddCommandItem2("SurFinSym32", -1, "", "SurFinSym 32", 21, "SurFinSym32", "EnableInDrw", mainItemID43, mTbO);

            cmdIndex50 = cmdGroup.AddCommandItem2("ExportPDF", -1, "", "Export PDF", 22, "ExportPDF", "EnableInDrw", mainItemID50, mTbO);
            cmdIndex51 = cmdGroup.AddCommandItem2("ExportTIF", -1, "", "Export TIF", 23, "ExportTIF", "EnableInDrw", mainItemID51, mTbO);
            cmdIndex52 = cmdGroup.AddCommandItem2("ExportDWG", -1, "", "Export DWG", 24, "ExportDWG", "EnableInDrw", mainItemID52, mTbO);
            cmdIndex53 = cmdGroup.AddCommandItem2("ExportSTP", -1, "", "Export STP", 25, "ExportSTP", "EnableInSld", mainItemID53, mTbO);

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();

            bool bResult;

            FlyoutGroup flyGroup = iCmdMgr.CreateFlyoutGroup(flyoutGroupID, "", "Flyout Group", "",
                                       cmdGroup.SmallMainIcon, cmdGroup.LargeMainIcon, cmdGroup.SmallIconList, cmdGroup.LargeIconList,
                                       "FlyoutCB1", "EnableInDrw");

            //flyGroup.AddCommandItem("FlyoutC01", "", 0, "FlyoutC01", "EnableInDrw");

            flyGroup.FlyoutType = (int)swCommandFlyoutStyle_e.swCommandFlyoutStyle_Simple;

            int[] docTypes = new int[]{//(int)swDocumentTypes_e.swDocPART,
                                           //(int)swDocumentTypes_e.swDocASSEMBLY,
                                           (int)swDocumentTypes_e.swDocDRAWING};
            foreach (int type in docTypes)
            {
                CommandTab cmdTab = iCmdMgr.GetCommandTab(type, Title);

                //if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.
                //Otherwise the ids won't matchup and the tab will be blank
                if (cmdTab != null & !getDataResult | ignorePrevious)
                {
                    bool res = iCmdMgr.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                //if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                if (cmdTab == null)
                {
                    cmdTab = iCmdMgr.AddCommandTab(type, Title);

                    #region CommandTabBox0
                    CommandTabBox cmdBox0 = cmdTab.AddCommandTabBox();

                    int[] cmdIDs = new int[4];
                    int[] TextType = new int[4];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex00);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex01);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex02);
                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[3] = cmdGroup.get_CommandID(cmdIndex03);
                    TextType[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    bResult = cmdBox0.AddCommands(cmdIDs, TextType);

                    #endregion

                    #region CommandTabBox1
                    CommandTabBox cmdBox1 = cmdTab.AddCommandTabBox();

                    cmdIDs = new int[9];
                    TextType = new int[9];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex10);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex12);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex13);
                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    cmdIDs[3] = cmdGroup.get_CommandID(cmdIndex14);
                    TextType[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    cmdIDs[4] = cmdGroup.get_CommandID(cmdIndex15);
                    TextType[4] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    cmdIDs[5] = cmdGroup.get_CommandID(cmdIndex16);
                    TextType[5] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    cmdIDs[6] = cmdGroup.get_CommandID(cmdIndex17);
                    TextType[6] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    cmdIDs[7] = cmdGroup.get_CommandID(cmdIndex18);
                    TextType[7] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    cmdIDs[8] = cmdGroup.get_CommandID(cmdIndex19);
                    TextType[8] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    bResult = cmdBox1.AddCommands(cmdIDs, TextType);
                    #endregion

                    cmdTab.AddSeparator(cmdBox1, cmdIDs[0]);

                    #region CommandTabBox2
                    CommandTabBox cmdBox2 = cmdTab.AddCommandTabBox();
                    cmdIDs = new int[4];
                    TextType = new int[4];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex20);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex21);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex22);
                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    cmdIDs[3] = cmdGroup.get_CommandID(cmdIndex23);
                    TextType[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    bResult = cmdBox2.AddCommands(cmdIDs, TextType);

                    #endregion

                    cmdTab.AddSeparator(cmdBox2, cmdIDs[0]);

                    #region CommandTabBox3
                    CommandTabBox cmdBox3 = cmdTab.AddCommandTabBox();
                    cmdIDs = new int[1];
                    TextType = new int[1];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex30);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    bResult = cmdBox3.AddCommands(cmdIDs, TextType);

                    #endregion

                    cmdTab.AddSeparator(cmdBox3, cmdIDs[0]);

                    #region CommandTabBox4
                    CommandTabBox cmdBox4 = cmdTab.AddCommandTabBox();
                    cmdIDs = new int[4];
                    TextType = new int[4];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex40);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex41);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex42);
                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    cmdIDs[3] = cmdGroup.get_CommandID(cmdIndex43);
                    TextType[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText;

                    bResult = cmdBox4.AddCommands(cmdIDs, TextType);

                    #endregion

                    cmdTab.AddSeparator(cmdBox4, cmdIDs[0]);

                    #region CommandTabBox5
                    CommandTabBox cmdBox5 = cmdTab.AddCommandTabBox();
                    cmdIDs = new int[4];
                    TextType = new int[4];

                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex50);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex51);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex52);
                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[3] = cmdGroup.get_CommandID(cmdIndex53);
                    TextType[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    bResult = cmdBox5.AddCommands(cmdIDs, TextType);

                    #endregion

                    cmdTab.AddSeparator(cmdBox5, cmdIDs[0]);

                    #region CommandTabBox6
                    CommandTabBox cmdBox6 = cmdTab.AddCommandTabBox();
                    cmdIDs = new int[1];
                    TextType = new int[1];

                    cmdIDs[0] = flyGroup.CmdID;
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow |
                                    (int)swCommandTabButtonFlyoutStyle_e.swCommandTabButton_ActionFlyout;

                    bResult = cmdBox6.AddCommands(cmdIDs, TextType);

                    #endregion

                    cmdTab.AddSeparator(cmdBox6, cmdIDs[0]);
                }
            }
            thisAssembly = null;
        }

        public void RemoveCommandMgr()
        {
            if(iBmpHad != null)
                iBmpHad.Dispose();

            iCmdMgr.RemoveCommandGroup(mainCmdGroupID);
            iCmdMgr.RemoveFlyoutGroup(flyoutGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }
            else
            {
                for (int i = 0; i < addinList.Count; i++)
                {
                    if (addinList[i] != storedList[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        #endregion

        #region UI Callbacks

        #region The decimal of dimension
        public void Decimal0()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);

                    iDisDim.SetPrecision3(0, (int)swDimensionPrecisionSettings_e.swDoNotChangePrecisionSetting,
                        (int)swDimensionPrecisionSettings_e.swDoNotChangePrecisionSetting,
                        (int)swDimensionPrecisionSettings_e.swDoNotChangePrecisionSetting);
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void Decimal1()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);

                    iDisDim.SetPrecision3(1, (int)swDimensionPrecisionSettings_e.swDoNotChangePrecisionSetting,
                        (int)swDimensionPrecisionSettings_e.swDoNotChangePrecisionSetting,
                        (int)swDimensionPrecisionSettings_e.swDoNotChangePrecisionSetting);
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void Decimal2()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);

                    iDisDim.SetPrecision3(2, (int)swDimensionPrecisionSettings_e.swDoNotChangePrecisionSetting,
                        (int)swDimensionPrecisionSettings_e.swDoNotChangePrecisionSetting,
                        (int)swDimensionPrecisionSettings_e.swDoNotChangePrecisionSetting);
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void Decimal3()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);

                    iDisDim.SetPrecision3(3, (int)swDimensionPrecisionSettings_e.swDoNotChangePrecisionSetting,
                        (int)swDimensionPrecisionSettings_e.swDoNotChangePrecisionSetting,
                        (int)swDimensionPrecisionSettings_e.swDoNotChangePrecisionSetting);
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        #endregion

        #region The number of dimension
        public void DimNumb0()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);

                    string textAbove = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove);
                    string textPrefix = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextPrefix);

                    if ((textAbove != null) && (textAbove != ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove, RemoveXx(textAbove, ""));
                    if ((textAbove == null) || (textAbove == ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextPrefix, RemoveXx(textPrefix, ""));
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void DimNumb2()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);

                    string textAbove = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove);
                    string textPrefix = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextPrefix);

                    if ((textAbove != null) && (textAbove != ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove, RemoveXx(textAbove, "2x"));
                    if ((textAbove == null) || (textAbove == ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextPrefix, RemoveXx(textPrefix, "2x"));
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void DimNumb3()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);

                    string textAbove = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove);
                    string textPrefix = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextPrefix);

                    if ((textAbove != null) && (textAbove != ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove, RemoveXx(textAbove, "3x"));
                    if ((textAbove == null) || (textAbove == ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextPrefix, RemoveXx(textPrefix, "3x"));
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void DimNumb4()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);

                    string textAbove = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove);
                    string textPrefix = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextPrefix);

                    if ((textAbove != null) && (textAbove != ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove, RemoveXx(textAbove, "4x"));
                    if ((textAbove == null) || (textAbove == ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextPrefix, RemoveXx(textPrefix, "4x"));
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void DimNumb5()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);

                    string textAbove = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove);
                    string textPrefix = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextPrefix);

                    if ((textAbove != null) && (textAbove != ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove, RemoveXx(textAbove, "5x"));
                    if ((textAbove == null) || (textAbove == ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextPrefix, RemoveXx(textPrefix, "5x"));
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void DimNumb6()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);

                    string textAbove = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove);
                    string textPrefix = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextPrefix);

                    if ((textAbove != null) && (textAbove != ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove, RemoveXx(textAbove, "6x"));
                    if ((textAbove == null) || (textAbove == ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextPrefix, RemoveXx(textPrefix, "6x"));
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void DimNumb7()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);

                    string textAbove = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove);
                    string textPrefix = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextPrefix);

                    if ((textAbove != null) && (textAbove != ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove, RemoveXx(textAbove, "7x"));
                    if ((textAbove == null) || (textAbove == ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextPrefix, RemoveXx(textPrefix, "7x"));
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void DimNumb8()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);

                    string textAbove = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove);
                    string textPrefix = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextPrefix);

                    if ((textAbove != null) && (textAbove != ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove, RemoveXx(textAbove, "8x"));
                    if ((textAbove == null) || (textAbove == ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextPrefix, RemoveXx(textPrefix, "8x"));
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void DimNumb9()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);

                    string textAbove = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove);
                    string textPrefix = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextPrefix);

                    if ((textAbove != null) && (textAbove != ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextCalloutAbove, RemoveXx(textAbove, "9x"));
                    if ((textAbove == null) || (textAbove == ""))
                        iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextPrefix, RemoveXx(textPrefix, "9x"));
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        #endregion

        #region The tolerance of dimension
        public void ToleranceN0()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                Dimension dim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);
                    dim = iDisDim.GetDimension2(0);

                    dim.SetToleranceType((int)swTolType_e.swTolNONE);
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void ToleranceH7()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                Dimension dim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);
                    dim = iDisDim.GetDimension2(0);

                    dim.SetToleranceType((int)swTolType_e.swTolFITWITHTOL);
                    dim.SetToleranceFitValues("", "H7");
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void Toleranceg6()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                Dimension dim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);
                    dim = iDisDim.GetDimension2(0);

                    dim.SetToleranceType((int)swTolType_e.swTolFITWITHTOL);
                    dim.SetToleranceFitValues("g6", "");
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void Tolerance01()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                Dimension dim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);
                    dim = iDisDim.GetDimension2(0);

                    dim.SetToleranceType((int)swTolType_e.swTolBILAT);
                    dim.SetToleranceValues(0, 0.00001);
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        #endregion

        #region The THR after the dimension
        public void ADDTHR()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            SelectionMgr selMgr = (SelectionMgr)modDoc2.SelectionManager;

            int selObjCount = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= selObjCount; i++)
            {
                IDisplayDimension iDisDim;
                if (selMgr.GetSelectedObjectType3(i, -1) == (int)swSelectType_e.swSelDIMENSIONS)
                {
                    iDisDim = (IDisplayDimension)selMgr.GetSelectedObject6(i, -1);

                    string textSuffix = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextSuffix);
                    //string textBelow = iDisDim.GetText((int)swDimensionTextParts_e.swDimensionTextCalloutBelow);

                    iDisDim.SetText((int)swDimensionTextParts_e.swDimensionTextSuffix, RemoveTHR(textSuffix, " THR"));
                }
            }

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        #endregion

        #region The Surface Finish Symbol
        public void SurFinSymN0()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            ModelDocExtension modelDocExt = modDoc2.Extension;
            DrawingDoc drwDoc = (DrawingDoc)iSwApp.IActiveDoc2;

            double[] point3D = (double[])drwDoc.GetInsertionPoint();
            modelDocExt.InsertSurfaceFinishSymbol3((int)swSFSymType_e.swSFDont_Machine, (int)swLeaderStyle_e.swNO_LEADER,
                                           point3D[0], point3D[1], point3D[2],
                                           (int)swSFLaySym_e.swSFNone, (int)swArrowStyle_e.swCLOSED_ARROWHEAD,
                                           "", "", "", "", "", "", "");

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void SurFinSym04()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            ModelDocExtension modelDocExt = modDoc2.Extension;
            DrawingDoc drwDoc = (DrawingDoc)iSwApp.IActiveDoc2;

            double[] point3D = (double[])drwDoc.GetInsertionPoint();
            modelDocExt.InsertSurfaceFinishSymbol3((int)swSFSymType_e.swSFMachining_Req, (int)swLeaderStyle_e.swNO_LEADER,
                                           point3D[0], point3D[1], point3D[2],
                                           (int)swSFLaySym_e.swSFNone, (int)swArrowStyle_e.swCLOSED_ARROWHEAD,
                                           "", "0.4", "", "Ra", "", "", "");

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void SurFinSym08()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            ModelDocExtension modelDocExt = modDoc2.Extension;
            DrawingDoc drwDoc = (DrawingDoc)iSwApp.IActiveDoc2;

            double[] point3D = (double[])drwDoc.GetInsertionPoint();
            modelDocExt.InsertSurfaceFinishSymbol3((int)swSFSymType_e.swSFMachining_Req, (int)swLeaderStyle_e.swNO_LEADER,
                                           point3D[0], point3D[1], point3D[2],
                                           (int)swSFLaySym_e.swSFNone, (int)swArrowStyle_e.swCLOSED_ARROWHEAD,
                                           "", "0.8", "", "Ra", "", "", "");

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void SurFinSym32()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            ModelDocExtension modelDocExt = modDoc2.Extension;
            DrawingDoc drwDoc = (DrawingDoc)iSwApp.IActiveDoc2;

            double[] point3D = (double[])drwDoc.GetInsertionPoint();
            modelDocExt.InsertSurfaceFinishSymbol3((int)swSFSymType_e.swSFMachining_Req, (int)swLeaderStyle_e.swNO_LEADER,
                                           point3D[0], point3D[1], point3D[2],
                                           (int)swSFLaySym_e.swSFNone, (int)swArrowStyle_e.swCLOSED_ARROWHEAD,
                                           "", "3.2", "", "Ra", "", "", "");

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        #endregion

        #region The Exporting files
        public void ExportPDF()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            ModelDocExtension modelDocExt = modDoc2.Extension;
            //DrawingDoc drwDoc = (DrawingDoc)iSwApp.IActiveDoc2;

            //ExportPdfData exportPDFData = iSwApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);
            //string sheetName = ((Sheet)drwDoc.GetCurrentSheet()).GetName();
            //exportPDFData.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportCurrentSheet, sheetName);
            //exportPDFData.ViewPdfAfterSaving = false;

            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportInColor, false);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportEmbedFonts, true);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportHighQuality, true);
            modelDocExt.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swPDFExportShadedDraftDPI,
                                                        (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified,
                                                        (int)swDetailingStandard_e.swDetailingStandardISO);
            modelDocExt.SetUserPreferenceInteger((int)swUserPreferenceIntegerValue_e.swPDFExportOleDPI,
                                                        (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified,
                                                        (int)swDetailingStandard_e.swDetailingStandardUserDefined);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportPrintHeaderFooter, false);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportUseCurrentPrintLineWeights, true);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportIncludeLayersNotToPrint, false);

            string filePath = Path.ChangeExtension(modDoc2.GetPathName(), "PDF");
            modelDocExt.SaveAs(filePath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                                   null, 0, 0);
            //modelDocExt.SaveAs(filePath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
            //                       exportPDFData, 0, 0);

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void ExportTIF()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            ModelDocExtension modelDocExt = modDoc2.Extension;

            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffImageType,
                                                        (int)swTiffImageType_e.swTiffImageBlackAndWhite);
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffCompressionScheme,
                                                        (int)swTiffCompressionScheme_e.swTiffUncompressed);
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swExportJpegCompression, 1);
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffScreenOrPrintCapture, 1);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swTiffPrintAllSheets, false);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swTiffPrintUseSheetSize, true);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swTiffPrintPadText, false);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swTIFIncludeLayersNotToPrint, false);
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffPrintDPI, 600);
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffPrintPaperSize,
                                                        (int)swDwgPaperSizes_e.swDwgPaperA4size);
            iSwApp.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperWidth, 297);
            iSwApp.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swTiffPrintDrawingPaperHeight, 210);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swTiffPrintScaleToFit, true);
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swTiffPrintScaleFactor, 1);

            string filePath = Path.ChangeExtension(modDoc2.GetPathName(), "TIF");
            modelDocExt.SaveAs(filePath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                                   null, 0, 0);

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void ExportDWG()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            ModelDocExtension modelDocExt = modDoc2.Extension;

            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDXFExportHiddenLayersOn, false);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDXFExportHiddenLayersWarnIsOn, true);
            //iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfUseSolidworksLayers, false);

            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion,
                                                        (int)swDxfFormat_e.swDxfFormat_R2000);
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputFonts, 0);
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputLineStyles, 0);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfMapping, false);
            //iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDXFDontShowMap, false);
            //iSwApp.SetUserPreferenceStringListValue((int)swUserPreferenceStringListValue_e.swDxfMappingFiles, "");
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputNoScale, 0);
            //iSwApp.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swDxfOutputScaleFactor, 1);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfEndPointMerge, false);
            //iSwApp.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swDxfMergingDistance, 0);
            //iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDXFHighQualityExport, false);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfExportSplinesAsSplines, true);
            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfMultiSheetOption,
                                                        (int)swDxfMultisheet_e.swDxfActiveSheetOnly);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfExportAllSheetsToPaperSpace, false);

            string filePath = Path.ChangeExtension(modDoc2.GetPathName(), "DWG");
            modelDocExt.SaveAs(filePath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                                   null, 0, 0);

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        public void ExportSTP()
        {
            ModelDoc2 modDoc2 = iSwApp.IActiveDoc2;
            ModelDocExtension modelDocExt = modDoc2.Extension;

            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swStepAP, 214);

            iSwApp.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swStepExportPreference,
                                                        (int)swAcisOutputGeometryPreference_e.swAcisOutputAsSolidAndSurface);
            //modelDocExt.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swStepExportConfigurationData, 
            //                                           (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, false);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swStepExportFaceEdgeProps, true);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swStepExportSplitPeriodic, true);
            iSwApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swStepExport3DCurveFeatures, false);
            modelDocExt.SetUserPreferenceString((int)swUserPreferenceStringValue_e.swDetailingDimensionStandardName,
                                                       (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified, null);

            string filePath = Path.ChangeExtension(modDoc2.GetPathName(), "STEP");
            modelDocExt.SaveAs(filePath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                                   null, 0, 0);

            modDoc2.WindowRedraw();
            //((ModelView)modDoc2.ActiveView).GraphicsRedraw(null);
        }

        #endregion

        #region UI FlyoutCallback Methods
        public void FlyoutCB1()
        {
            FlyoutGroup flyGroup = iCmdMgr.GetFlyoutGroup(flyoutGroupID);
            flyGroup.RemoveAllCommandItems();

            flyGroup.AddCommandItem(System.DateTime.Now.ToLongTimeString(), "", 0, "FlyoutC01", "EnableInDrw");
        }
        
        public void FlyoutC01()
        {
            iSwApp.SendMsgToUser("C01");
        }

        #endregion

        #region UI Callback Methods Enable
        public int EnableInDrw()
        {
            if ((iSwApp.ActiveDoc != null) && (iSwApp.IActiveDoc2.GetType() == (int)swDocumentTypes_e.swDocDRAWING))
                return 1;
            else
                return 0;
        }

        public int EnableInPrt()
        {
            if ((iSwApp.ActiveDoc != null) && (iSwApp.IActiveDoc2.GetType() == (int)swDocumentTypes_e.swDocPART))
                return 1;
            else
                return 0;
        }

        public int EnableInAsm()
        {
            if ((iSwApp.ActiveDoc != null) && (iSwApp.IActiveDoc2.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY))
                return 1;
            else
                return 0;
        }

        public int EnableInSld()
        {
            if ((iSwApp.ActiveDoc != null) && (iSwApp.IActiveDoc2.GetType() == (int)swDocumentTypes_e.swDocPART))
                return 1;
            else if ((iSwApp.ActiveDoc != null) && (iSwApp.IActiveDoc2.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY))
                return 1;
            else
                return 0;
        }

        #endregion

        #endregion

        #region Private Programs Methodes
        private string RemoveXx(string str, string prefix)
        {
            if (str == null && prefix == null)
                return null;
            if (str == null || str.Length <= 1 || (!char.IsDigit(str[0])))
                return prefix + str;

            for (int counter = 1; counter < str.Length; counter++)
            {
                if ((str[counter] == 'x') || (str[counter] == 'X') || (str[counter] == '-'))
                {
                    return prefix + str.PadLeft(counter + 1).Remove(0, counter + 1);
                }
                else if (!char.IsDigit(str[0]))
                {
                    return prefix + str;
                }
            }

            return prefix + str;
        }

        private string RemoveTHR(string str, string suffix)
        {
            if (str == null && suffix == null)
                return null;
            if (str == null || str.Length < 3)
                return str + suffix;

            string strUp = str.ToUpper();

            if (strUp.Substring(strUp.Length - 3) == "THR")
            {
                if ((strUp.Length >= 4) && (strUp.Substring(strUp.Length - 4) == " THR"))
                {
                    str = str.Remove(str.Length - 4, 4);
                    return str + suffix;
                }
                if ((strUp.Length >= 4) && (strUp.Substring(strUp.Length - 4) == "-THR"))
                {
                    str = str.Remove(str.Length - 4, 4);
                    return str + suffix;
                }

                str = str.Remove(str.Length - 3, 3);
                return str + suffix;
            }

            return str + suffix;
        }

        #endregion
    }
}