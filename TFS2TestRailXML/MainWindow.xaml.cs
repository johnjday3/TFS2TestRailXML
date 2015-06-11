using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Xml;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;

namespace TFS2TestRailXML
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private TfsTeamProjectCollection _tfs;
        ITestManagementTeamProject _testProject;
        ITestPlanCollection _testPlanCollection;
        private ITestSuiteBase _suite;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

        }

        void BtnConnect_Click(object sender, RoutedEventArgs e)
        {            
            _tfs = null;
            TbTfs.Text = null;
            var tpp = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false);
            tpp.ShowDialog();

            if (tpp.SelectedTeamProjectCollection == null) return;

            _tfs = tpp.SelectedTeamProjectCollection;
                
            var testService = (ITestManagementService)_tfs.GetService(typeof(ITestManagementService));

            _testProject = testService.GetTeamProject(tpp.SelectedProjects[0].Name);

            TbTfs.Text = _tfs.Name + "\\" + _testProject;

            _testPlanCollection = _testProject.TestPlans.Query("Select * from TestPlan");

            ProjectSelected_GetTestPlans();

            
            btnGenerateXMLFile.IsEnabled = false;
        }

        private void BtnOpenFileDialog_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog1 = new System.Windows.Forms.SaveFileDialog
            {
                InitialDirectory = Environment.SpecialFolder.MyDocuments.ToString(),
                Filter = Properties.Resources.MainWindow_BtnOpenFileDialog_Click,
                FilterIndex = 1
            };

            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                TbFileName.Text = saveFileDialog1.FileName;
            }
            else
            {
                MessageBox.Show("Please choose a valid filename");
            }
        }

        private void ProjectSelected_GetTestPlans()
        {
            CbTestPlans.ItemsSource = _testPlanCollection;
            CbTestPlans.DisplayMemberPath = NameProperty.ToString();
        }

        private void CbTestPlans_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TvSuites.Items.Clear();
            AddTestSuites(CbTestPlans.SelectedItem as ITestPlan);
        }

        private void TvSuites_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            btnGenerateXMLFile.IsEnabled = true;
        }

        void AddTestSuites(ITestPlan selectedTestPlan)
        // Adds the selected Test Plan to the tree view as the root
        // then gets all the sub suites
        {
            if (selectedTestPlan == null) return;
            var root = new TreeViewItem {Header = selectedTestPlan.RootSuite.Title};

            TvSuites.Items.Add(root);
            root.Tag = selectedTestPlan.RootSuite.Id;
                
            AddSubSuites(selectedTestPlan.RootSuite.SubSuites, root);
        }

        static void AddSubSuites(IEnumerable<ITestSuiteBase> subSuiteEntries, ItemsControl treeNode)
        // Recursively adds all the sub suites to the tree view
        {
            foreach (var suite in subSuiteEntries)
            {
                var suiteTree = new TreeViewItem {Header = suite.Title};
                treeNode.Items.Add(suiteTree);
                suiteTree.Tag = suite.Id;
                if (suite.TestSuiteType == TestSuiteType.StaticTestSuite)
                {
                    var suite1 = suite as IStaticTestSuite;
                    if (suite1.SubSuites.Count > 0)
                    {
                        AddSubSuites(suite1.SubSuites, suiteTree);
                    }
                }
            }
        }

        void BtnGenerateXMLFile_Click(object sender, RoutedEventArgs e)
        {
            if (TbFileName.Text == null || TbFileName.Text.Length.Equals(0))
            {
                MessageBox.Show("Please Enter a valid file path");
            }
            else
            {
                if (TvSuites.SelectedValue != null)
                {
                    var tvItem = TvSuites.SelectedItem as TreeViewItem;

                    _suite = _testProject.TestSuites.Find(Convert.ToInt32(tvItem.Tag.ToString()));

                    if (_suite != null)
                        WriteTestPlanToXml(_suite);
                }
                else
                {
                    MessageBox.Show("Please select a test suite");
                }
            }
        }

        void WriteTestPlanToXml(ITestSuiteBase rootSuite)
        {
            try
            {
                var xmlDoc = new XmlDocument();

                if (rootSuite.TestSuiteType == TestSuiteType.StaticTestSuite)
                {
                    WriteRootSuite(rootSuite as IStaticTestSuite, xmlDoc);                    
                }
                if (rootSuite.TestSuiteType == TestSuiteType.DynamicTestSuite)
                {
                    WriteRootSuite(rootSuite as IDynamicTestSuite, xmlDoc);
                }
                if (rootSuite.TestSuiteType == TestSuiteType.RequirementTestSuite)
                {
                    WriteRootSuite(rootSuite as IRequirementTestSuite, xmlDoc);
                }
                

                xmlDoc.Save(TbFileName.Text);

                MessageBox.Show("File has been saved at " + TbFileName.Text);
            }
            catch (Exception theException)
            {
                var errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        void WriteRootSuite(IStaticTestSuite testSuite, XmlDocument xmlDoc)
        {
            XmlNode rootNode = xmlDoc.CreateElement("suite");
            xmlDoc.AppendChild(rootNode);

            XmlNode idNode = xmlDoc.CreateElement("id");
            rootNode.AppendChild(idNode);

            XmlNode rootnameNode = xmlDoc.CreateElement("name");
            rootnameNode.InnerText = testSuite.Plan.Name;
            rootNode.AppendChild(rootnameNode);

            XmlNode rootdescNode = xmlDoc.CreateElement("description");
            rootdescNode.InnerText = testSuite.Plan.Description;
            rootNode.AppendChild(rootdescNode);

            XmlNode sectionsNode = xmlDoc.CreateElement("sections");
            rootNode.AppendChild(sectionsNode);

            if (testSuite.TestSuiteType == TestSuiteType.StaticTestSuite)
            {
                if (testSuite.Entries.Any(suiteEntry => suiteEntry.EntryType == TestSuiteEntryType.TestCase))
                {
                    XmlNode sectionNode = xmlDoc.CreateElement("section");

                    XmlNode nameNode = xmlDoc.CreateElement("name");
                    sectionNode.AppendChild(nameNode);

                    nameNode.InnerText = "All Test Cases";

                    XmlNode descNode = xmlDoc.CreateElement("description");
                    sectionNode.AppendChild(descNode);
                    XmlNode casesNode = xmlDoc.CreateElement("cases");
                    sectionNode.AppendChild(casesNode);

                    foreach (var suiteEntry in testSuite.Entries.Where(suiteEntry => suiteEntry.EntryType == TestSuiteEntryType.TestCase))
                    {
                        var caseNode = WriteTestCase(suiteEntry.TestCase, xmlDoc);
                        casesNode.AppendChild(caseNode);
                    }
                    sectionsNode.AppendChild(sectionNode);
                }

                foreach (var suiteEntry in testSuite.Entries.Where(suiteEntry => suiteEntry.EntryType == TestSuiteEntryType.StaticTestSuite ||
                                                                                   suiteEntry.EntryType == TestSuiteEntryType.RequirementTestSuite ||
                                                                                   suiteEntry.EntryType == TestSuiteEntryType.DynamicTestSuite))
                {
                    if (suiteEntry.EntryType == TestSuiteEntryType.StaticTestSuite)
                    {
                        var suite = suiteEntry.TestSuite as IStaticTestSuite;

                        var subSection = GetTestSuites(suite, xmlDoc);
                        sectionsNode.AppendChild(subSection);

                    }
                    if (suiteEntry.EntryType == TestSuiteEntryType.RequirementTestSuite)
                    {
                        var suite = suiteEntry.TestSuite as IRequirementTestSuite;
                        GetTestCases(suite, xmlDoc, sectionsNode);
                    }
                    if (suiteEntry.EntryType == TestSuiteEntryType.DynamicTestSuite)
                    {
                        var suite = suiteEntry.TestSuite as IDynamicTestSuite;
                        GetTestCases(suite, xmlDoc, sectionsNode);
                    }
                }
            }
            /*if (testSuite.TestSuiteType == TestSuiteType.DynamicTestSuite)
            {
                GetTestCases(testSuite as IDynamicTestSuite, xmlDoc, sectionsNode);
            }
            if (testSuite.TestSuiteType == TestSuiteType.RequirementTestSuite)
            {
                GetTestCases(testSuite as IRequirementTestSuite, xmlDoc, sectionsNode);
            }*/
        }

        void WriteRootSuite(IDynamicTestSuite testSuite, XmlDocument xmlDoc)
        {
            XmlNode rootNode = xmlDoc.CreateElement("suite");
            xmlDoc.AppendChild(rootNode);

            XmlNode idNode = xmlDoc.CreateElement("id");
            rootNode.AppendChild(idNode);

            XmlNode rootnameNode = xmlDoc.CreateElement("name");
            rootnameNode.InnerText = testSuite.Plan.Name;
            rootNode.AppendChild(rootnameNode);

            XmlNode rootdescNode = xmlDoc.CreateElement("description");
            rootdescNode.InnerText = testSuite.Plan.Description;
            rootNode.AppendChild(rootdescNode);

            XmlNode sectionsNode = xmlDoc.CreateElement("sections");
            rootNode.AppendChild(sectionsNode);
            GetTestCases(testSuite, xmlDoc, sectionsNode);    
        }

        void WriteRootSuite(IRequirementTestSuite testSuite, XmlDocument xmlDoc)
        {
            XmlNode rootNode = xmlDoc.CreateElement("suite");
            xmlDoc.AppendChild(rootNode);

            XmlNode idNode = xmlDoc.CreateElement("id");
            rootNode.AppendChild(idNode);

            XmlNode rootnameNode = xmlDoc.CreateElement("name");
            rootnameNode.InnerText = testSuite.Plan.Name;
            rootNode.AppendChild(rootnameNode);

            XmlNode rootdescNode = xmlDoc.CreateElement("description");
            rootdescNode.InnerText = testSuite.Plan.Description;
            rootNode.AppendChild(rootdescNode);

            XmlNode sectionsNode = xmlDoc.CreateElement("sections");
            rootNode.AppendChild(sectionsNode);
            GetTestCases(testSuite, xmlDoc, sectionsNode);
        }

        private XmlNode GetTestSuites(IStaticTestSuite staticTestSuite, XmlDocument xmlDoc)
        {
            XmlNode sectionNode = xmlDoc.CreateElement("section");

            XmlNode nameNode = xmlDoc.CreateElement("name");
            sectionNode.AppendChild(nameNode);
            nameNode.InnerText = staticTestSuite.Title;

            XmlNode descNode = xmlDoc.CreateElement("description");
            sectionNode.AppendChild(descNode);

            if (staticTestSuite.Entries.Any(suiteEntry => suiteEntry.EntryType == TestSuiteEntryType.TestCase))
            {
                XmlNode casesNode = xmlDoc.CreateElement("cases");
                sectionNode.AppendChild(casesNode);

                foreach (
                    var suiteEntry in
                        staticTestSuite.Entries.Where(suiteEntry => suiteEntry.EntryType == TestSuiteEntryType.TestCase)
                    )
                {
                    var caseNode = WriteTestCase(suiteEntry.TestCase, xmlDoc);
                    casesNode.AppendChild(caseNode);
                }
            }

            if (staticTestSuite.SubSuites.Count > 0)
            {
                XmlNode sectionsNode = xmlDoc.CreateElement("sections");

                foreach (
                    var suiteEntry in
                        staticTestSuite.Entries.Where(
                            suiteEntry => suiteEntry.EntryType == TestSuiteEntryType.StaticTestSuite ||
                                          suiteEntry.EntryType == TestSuiteEntryType.RequirementTestSuite ||
                                          suiteEntry.EntryType == TestSuiteEntryType.DynamicTestSuite))
                {
                    if (suiteEntry.EntryType == TestSuiteEntryType.StaticTestSuite)
                    {
                        var suite = suiteEntry.TestSuite as IStaticTestSuite;

                        var subSection = GetTestSuites(suite, xmlDoc);
                        sectionsNode.AppendChild(subSection);

                    }
                    if (suiteEntry.EntryType == TestSuiteEntryType.RequirementTestSuite)
                    {
                        var suite = suiteEntry.TestSuite as IRequirementTestSuite;
                        GetTestCases(suite, xmlDoc, sectionsNode);
                    }
                    if (suiteEntry.EntryType == TestSuiteEntryType.DynamicTestSuite)
                    {
                        var suite = suiteEntry.TestSuite as IDynamicTestSuite;
                        GetTestCases(suite, xmlDoc, sectionsNode);
                    }
                }

                sectionNode.AppendChild(sectionsNode);
            }
            return sectionNode;
        }


        private void GetTestCases(IRequirementTestSuite requirementTestSuite, XmlDocument xmlDoc, XmlNode node) //sectionsNode has to be "sections"
        {
            var casesNode = WriteSuite(requirementTestSuite, xmlDoc, node); //WriteSuite sectionsNode has to be "sections"
            foreach (var testCase in requirementTestSuite.AllTestCases)
            {
                var caseNode = WriteTestCase(testCase, xmlDoc);
                casesNode.AppendChild(caseNode);
            }
        }

        private void GetTestCases(IDynamicTestSuite dynamicTestSuite, XmlDocument xmlDoc, XmlNode node) //sectionsNode has to be "sections"
        {
            var casesNode = WriteSuite(dynamicTestSuite, xmlDoc, node);
            
            foreach (var testCase in dynamicTestSuite.AllTestCases)
            {
                var caseNode = WriteTestCase(testCase, xmlDoc);
                casesNode.AppendChild(caseNode);
            }
        }

       XmlNode WriteSuite(IRequirementTestSuite testSuite, XmlDocument xmlDoc, XmlNode node) //sectionsNode has to be "sections"
        {
            XmlNode sectionNode = xmlDoc.CreateElement("section");
            node.AppendChild(sectionNode);
            XmlNode nameNode = xmlDoc.CreateElement("name");
            nameNode.InnerText = testSuite.Title;
            sectionNode.AppendChild(nameNode);
            XmlNode descNode = xmlDoc.CreateElement("description");
            sectionNode.AppendChild(descNode);
            XmlNode casesNode = xmlDoc.CreateElement("cases");
            sectionNode.AppendChild(casesNode);

            return casesNode;
        }

        XmlNode WriteSuite(IDynamicTestSuite testSuite, XmlDocument xmlDoc, XmlNode node) //sectionsNode has to be "sections"
        {
            XmlNode sectionNode = xmlDoc.CreateElement("section");
            node.AppendChild(sectionNode);
            XmlNode nameNode = xmlDoc.CreateElement("name");
            nameNode.InnerText = testSuite.Title;
            sectionNode.AppendChild(nameNode);
            XmlNode descNode = xmlDoc.CreateElement("description");
            sectionNode.AppendChild(descNode);
            XmlNode casesNode = xmlDoc.CreateElement("cases");
            sectionNode.AppendChild(casesNode);

            return casesNode;
        }

        XmlNode WriteTestCase(ITestCase testCase, XmlDocument xmlDoc) //sectionsNode has to be "cases"
        {
            XmlNode caseNode = xmlDoc.CreateElement("case");
            XmlNode idNode = xmlDoc.CreateElement("id");
            caseNode.AppendChild(idNode);
            XmlNode titleNode = xmlDoc.CreateElement("title");
            titleNode.InnerText = testCase.Title;
            caseNode.AppendChild(titleNode);
            XmlNode typeNode = xmlDoc.CreateElement("type");
            caseNode.AppendChild(typeNode);
            XmlNode priorityNode = xmlDoc.CreateElement("priority");
            caseNode.AppendChild(priorityNode);
            XmlNode estimateNode = xmlDoc.CreateElement("estimate");
            caseNode.AppendChild(estimateNode);
            XmlNode milestoneNode = xmlDoc.CreateElement("milestone");
            caseNode.AppendChild(milestoneNode);
            XmlNode referencesNode = xmlDoc.CreateElement("references");
            caseNode.AppendChild(referencesNode);
            XmlNode customNode = xmlDoc.CreateElement("custom");
            caseNode.AppendChild(customNode);
            XmlNode stepsNode = xmlDoc.CreateElement("steps_separated");
            customNode.AppendChild(stepsNode);
           
            var i = 1;

            foreach (ITestAction action in testCase.Actions)
            {
                var sharedRef = action as ISharedStepReference;

                if (sharedRef != null)
                {
                    var sharedStep = sharedRef.FindSharedStep();
                    foreach (var testStep in sharedStep.Actions.Select(sharedAction => sharedAction as ITestStep))
                    {
                        WriteTestSteps(i, testStep.Title.ToString(), testStep.ExpectedResult.ToString(), xmlDoc, stepsNode);
                        i++;
                    }
                }
                else
                {
                    var testStep = action as ITestStep;
                    WriteTestSteps(i, testStep.Title.ToString(), testStep.ExpectedResult.ToString(), xmlDoc, stepsNode);
                    i++;
                }
            } //end of foreach test action
            return caseNode;
        }
        
        void WriteTestSteps(int i, string testAction, string expectedResult, XmlDocument xmlDoc, XmlNode node)
        {
            XmlNode stepNode = xmlDoc.CreateElement("step");
            node.AppendChild(stepNode);
            XmlNode indexNode = xmlDoc.CreateElement("index");
            indexNode.InnerText = i.ToString();
            stepNode.AppendChild(indexNode);
            XmlNode contentNode = xmlDoc.CreateElement("content");
            contentNode.InnerText = testAction;
            stepNode.AppendChild(contentNode);
            XmlNode expectedNode = xmlDoc.CreateElement("expected");
            expectedNode.InnerText = expectedResult;
            stepNode.AppendChild(expectedNode);
        }
    }
}