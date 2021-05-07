using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;

namespace KSP.Forms
{
    public partial class JoshForm01 : System.Windows.Forms.Form
    {
        ButtonEE7Parameter myEE7Parameter;
        ExternalEvent myEE7ActionParameter;
        public Document doc { get; set; }

        LinePatternsWeightsFalse mylinePatternsWeightsFalse;
        ExternalEvent myMakeLinePatterns;

        LinePatternsWeightsTrue mylinePatternsWeightsTrue;
        ExternalEvent myMakeLineWeights;

        public JoshForm01()
        {
            InitializeComponent();

            myEE7Parameter = new ButtonEE7Parameter();
            myEE7ActionParameter = ExternalEvent.Create(myEE7Parameter);

            mylinePatternsWeightsFalse = new LinePatternsWeightsFalse();
            myMakeLinePatterns = ExternalEvent.Create(mylinePatternsWeightsFalse);

            mylinePatternsWeightsTrue = new LinePatternsWeightsTrue();
            myMakeLineWeights = ExternalEvent.Create(mylinePatternsWeightsTrue);
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            myMakeLineWeights.Raise();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            myMakeLinePatterns.Raise();
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            //throw new InvalidOperationException();

            try
            {
                //throw new InvalidOperationException();
                //using (Transaction t = new Transaction(doc, "Set a parameters"))
                //{
                //    t.Start();
                //        doc.ProjectInformation.GetParameters("Project Name")[0].Set("Новый проект");  //this needs to change in two places
                //    t.Commit();
                //}


                myEE7ActionParameter.Raise();  				
            }

            #region catch and finally
            catch (Exception ex)
            {
                TaskDialog.Show("Catch", "Failed due to:" + Environment.NewLine + ex.Message);
            }
            finally
            {

            }
            #endregion			
        }

        public class LinePatternsWeightsTrue : IExternalEventHandler
        {

            public void Execute(UIApplication a)
            {
                UIDocument uidoc = a.ActiveUIDocument;
                Document doc = uidoc.Document;

                Transaction transaction = new Transaction(doc);
                transaction.Start("Draw Line Patterns or Weights");

                DrawLines myThis = new DrawLines();

                myThis._99_DrawLinePatterns(true, false, uidoc);
                transaction.Commit();
                return;
            }

            public string GetName()
            {
                return "External Event Example";
            }
        }

        [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
        [Autodesk.Revit.DB.Macros.AddInId("23CF5F71-5468-438D-97C7-554F4F782936")]
        public class LinePatternsWeightsFalse : IExternalEventHandler
        {

            public void Execute(UIApplication a)
            {

                UIDocument uidoc = a.ActiveUIDocument;
                Document doc = uidoc.Document;

                Transaction transaction = new Transaction(doc);
                transaction.Start("Draw Line Patterns or Weights");

                DrawLines myThis = new DrawLines();

                myThis._99_DrawLinePatterns(false, false, uidoc);

                transaction.Commit();
                return;
            }


            public string GetName()
            {
                return "External Event Example";
            }
        }
    }
}
