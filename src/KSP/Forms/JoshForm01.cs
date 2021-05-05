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
        public Document doc { get; set; }

        linePatternsWeightsFalse mylinePatternsWeightsFalse;
        ExternalEvent myMakeLinePatterns;

        linePatternsWeightsTrue mylinePatternsWeightsTrue;
        ExternalEvent myMakeLineWeights;

        public JoshForm01()
        {
            InitializeComponent();
            mylinePatternsWeightsFalse = new linePatternsWeightsFalse();
            myMakeLinePatterns = ExternalEvent.Create(mylinePatternsWeightsFalse);

            mylinePatternsWeightsTrue = new linePatternsWeightsTrue();
            myMakeLineWeights = ExternalEvent.Create(mylinePatternsWeightsTrue);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            myMakeLineWeights.Raise();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            myMakeLinePatterns.Raise();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            throw new InvalidOperationException();

            try
            {
                throw new InvalidOperationException();
                using (Transaction t = new Transaction(doc, "Set a parameters"))
                {
                    t.Start();
                    doc.ProjectInformation.GetParameters("Project Name")[0].Set("Space Elevator");  //this needs to change in two places
                    t.Commit();
                }


                //myEE7ActionParameter.Raise();  				
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

        public class linePatternsWeightsTrue : IExternalEventHandler
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
        public class linePatternsWeightsFalse : IExternalEventHandler
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
