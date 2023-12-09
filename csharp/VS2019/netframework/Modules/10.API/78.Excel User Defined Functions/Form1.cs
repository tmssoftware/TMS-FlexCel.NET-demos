using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Text;

using FlexCel.Render;

namespace ExcelUserDefinedFunctions
{
    /// <summary>
    /// An example on how to recalculate user defined functions.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
        }

        private string PathToExe
        {
            get
            {
                return Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;
            }
        }

        /// <summary>
        /// Loads the user defined functions into the Excel recalculating engine.
        /// </summary>
        /// <param name="Xls"></param>
        private void LoadUdfs(ExcelFile Xls)
        {
            Xls.AddUserDefinedFunction(TUserDefinedFunctionScope.Local, TUserDefinedFunctionLocation.Internal, new SumCellsWithSameColor());
            Xls.AddUserDefinedFunction(TUserDefinedFunctionScope.Local, TUserDefinedFunctionLocation.Internal, new IsPrime());
            Xls.AddUserDefinedFunction(TUserDefinedFunctionScope.Local, TUserDefinedFunctionLocation.Internal, new BoolChoose());
            Xls.AddUserDefinedFunction(TUserDefinedFunctionScope.Local, TUserDefinedFunctionLocation.Internal, new Lowest());
        }

        private void AddData(ExcelFile Xls)
        {
            LoadUdfs(Xls); //Register our custom functions. As we are using a local scope, we need to register them each time.

            Xls.Open(Path.Combine(PathToExe, "udfs.xls"));  //Open the file we want to manipulate.

            //Fill the cell range with other values so we can see how the sheet is recalculated by FlexCel.
            TXlsCellRange Data = Xls.GetNamedRange("Data", -1);
            for (int r = Data.Top; r < Data.Bottom; r++)
            {
                Xls.SetCellValue(r, Data.Left, r - Data.Top);
            }

            //Add an UDF to the sheet. We can enter the fucntion "BoolChoose" here because it was registered into FlexCel in LoadUDF()
            //If it hadn't been registered, this line would raise an Exception of an unknown function.
            string FmlaText = "=BoolChoose(TRUE,\"This formula was entered with FlexCel!\",\"It shouldn't display this\")";
            Xls.SetCellValue(11, 1, new TFormula(FmlaText));

            //Verify the UDF entered is correct. We can read any udf from Excel, even if it is not registered with AddUserDefinedFunction.
            object o = Xls.GetCellValue(11, 1);
            TFormula fm = o as TFormula;
            Debug.Assert(fm != null, "The cell must contain a formula");
            if (fm != null) Debug.Assert(fm.Text == FmlaText, "Error in Formula: It should be \"" + FmlaText + "\" and it is \"" + fm.Text + "\"");

            //Recalc the sheet. As we are not saving it yet, we ned to make a manual recalc.
            Xls.Recalc();

            //Export the file to PDF so we can see the values calculated by FlexCel without Excel recalculating them.
            using (FlexCelPdfExport pdf = new FlexCelPdfExport(Xls, true))
            {
                pdf.Export(saveFileDialog1.FileName);
            }

            //Save the file as xls too so we can compare.
            Xls.Save(Path.ChangeExtension(saveFileDialog1.FileName, "xls"));
        }


        private void button1_Click(object sender, System.EventArgs e)
        {
            AutoRun();
        }

        public void AutoRun()
        {
            if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;
            ExcelFile Xls = new XlsFile(true);
            AddData(Xls);
            if (MessageBox.Show("Do you want to open the generated files (PDF and XLS)?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                using (Process p = new Process())
                {               
                    p.StartInfo.FileName = saveFileDialog1.FileName;
                    p.StartInfo.UseShellExecute = true;
                    p.Start();
                }
                using (Process p = new Process())
                {               
                    p.StartInfo.FileName = Path.ChangeExtension(saveFileDialog1.FileName, "xls");
                    p.StartInfo.UseShellExecute = true;
                    p.Start();
                }
            }

        }

        /// <summary>
        /// This is the method that will be called by the ASP.NET front end. It returns an array of bytes 
        /// with the report data, so the ASP.NET application can stream it to the client.
        /// </summary>
        /// <returns>The generated file as a byte array.</returns>
        public byte[] WebRun()
        {
            ExcelFile Xls = new XlsFile(true);
            AddData(Xls);

            using (MemoryStream OutStream = new MemoryStream())
            {
                Xls.Save(OutStream);
                return OutStream.ToArray();
            }
        }
    }

    #region UDF definitions
    /// <summary>
    /// Implements a custom function that will sum the cells in a range that have the same
    /// color of the source cell. This function mimics the VBA macro in the example, so when
    /// recalculating the sheet with FlexCel you will get the same results as with Excel.
    /// </summary>
    public class SumCellsWithSameColor: TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new instance and registers the class in the FlexCel recalculating engine as "SumCellsWithSameColor".
        /// </summary>
        public SumCellsWithSameColor() : base("SumCellsWithSameColor")
        {
        }

        /// <summary>
        /// Returns the sum of cells in a range that have the same color as a reference cell.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters">In this case we expect 2 parameters, first the reference cell and then
        /// the range in which to sum. We will return an error otherwise.</param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            #region Get Parameters
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 2, out Err)) return Err;

            //The first parameter should be a range
            TXls3DRange SourceCell;
            if (!TryGetCellRange(parameters[0], out SourceCell, out Err)) return Err;

            //The second parameter should be a range too.
            TXls3DRange SumRange;
            if (!TryGetCellRange(parameters[1], out SumRange, out Err)) return Err;
            #endregion

            //Get the color in SourceCell. Note that if Source cell is a range with more than one cell,
            //we will use the first cell in the range. Also, as different colors can have the same rgb value, we will compare the actual RGB values, not the ExcelColors
            TFlxFormat fmt = arguments.Xls.GetCellVisibleFormatDef(SourceCell.Sheet1, SourceCell.Top, SourceCell.Left);
            int SourceColor = fmt.FillPattern.FgColor.ToColor(arguments.Xls).ToArgb();

            double Result = 0;
            //Loop in the sum range and sum the corresponding values.
            for (int s = SumRange.Sheet1; s <= SumRange.Sheet2; s++)
            {
                for (int r = SumRange.Top; r <= SumRange.Bottom; r++)
                {
                    for (int c = SumRange.Left; c <= SumRange.Right; c++)
                    {
                        int XF = -1;
                        object val = arguments.Xls.GetCellValue(s, r, c, ref XF);
                        if (val is double) //we will only sum numeric values.
                        {
                            TFlxFormat sumfmt = arguments.Xls.GetCellVisibleFormatDef(s, r, c);
                            if (sumfmt.FillPattern.FgColor.ToColor(arguments.Xls).ToArgb() == SourceColor)
                            {
                                Result += (double)val;
                            }
                        }
                    }
                }
            }
            return Result;
        }
    }


    /// <summary>
    /// Implements a custom function that will return true if a number is prime.
    /// This function mimics the VBA macro in the example, so when
    /// recalculating the sheet with FlexCel you will get the same results as with Excel.
    /// </summary>
    public class IsPrime: TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new instance and registers the class in the FlexCel recalculating engine as "IsPrime".
        /// </summary>
        public IsPrime() : base("IsPrime")
        {
        }

        /// <summary>
        /// Returns true if a number is prime.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters">In this case we expect 1 parameter with the number. We will return an error otherwise.</param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            #region Get Parameters
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 1, out Err)) return Err;

            //The parameter should be a double or a range.
            double Number;
            if (!TryGetDouble(arguments.Xls, parameters[0], out Number, out Err)) return Err;
            #endregion

            //Return true if the number is prime.
            int n = Convert.ToInt32(Number);
            if (n == 2) return true;
            if (n < 2 || n % 2 == 0) return false;
            for (int i = 3; i <= Math.Sqrt(n); i += 2)
            {
                if (n % i == 0) return false;
            }
            return true;
        }
    }

    /// <summary>
    /// Implements a custom function that will choose between two different strings.
    /// This function mimics the VBA macro in the example, so when
    /// recalculating the sheet with FlexCel you will get the same results as with Excel.
    /// </summary>
    public class BoolChoose: TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new instance and registers the class in the FlexCel recalculating engine as "BoolChoose".
        /// </summary>
        public BoolChoose() : base("BoolChoose")
        {
        }

        /// <summary>
        /// Chooses between 2 different strings.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters">In this case we expect 3 parameters: The first is a boolean, and the other 2 strings. We will return an error otherwise.</param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            #region Get Parameters
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 3, out Err)) return Err;

            //The first parameter should be a boolean.
            bool ChooseFirst;
            if (!TryGetBoolean(arguments.Xls, parameters[0], out ChooseFirst, out Err)) return Err;

            //The second parameter should be a string.
            string s1;
            if (!TryGetString(arguments.Xls, parameters[1], out s1, out Err)) return Err;

            //The third parameter should be a string.
            string s2;
            if (!TryGetString(arguments.Xls, parameters[2], out s2, out Err)) return Err;
            #endregion

            //Return s1 or s2 depending on ChooseFirst
            if (ChooseFirst) return s1; else return s2;
        }
    }

    /// <summary>
    /// Implements a custom function that will choose the lowest member in an array.
    /// This function mimics the VBA macro in the example, so when
    /// recalculating the sheet with FlexCel you will get the same results as with Excel.
    /// </summary>
    public class Lowest: TUserDefinedFunction
    {
        /// <summary>
        /// Creates a new instance and registers the class in the FlexCel recalculating engine as "Lowest".
        /// </summary>
        public Lowest() : base("Lowest")
        {
        }

        /// <summary>
        /// Chooses the lowest element in an array.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="parameters">In this case we expect 1 parameter that should be an array. We will return an error otherwise.</param>
        /// <returns></returns>
        public override object Evaluate(TUdfEventArgs arguments, object[] parameters)
        {
            #region Get Parameters
            TFlxFormulaErrorValue Err;
            if (!CheckParameters(parameters, 1, out Err)) return Err;

            //The first parameter should be an array.
            object[,] SourceArray;
            if (!TryGetArray(arguments.Xls, parameters[0], out SourceArray, out Err)) return Err;
            #endregion

            double Result = 0;
            bool First = true;
            foreach (object o in SourceArray)
            {
                if (o is double)
                {
                    if (First)
                    {
                        First = false;
                        Result = (double)o;
                    }
                    else
                    {
                        if ((double)o < Result) Result = (double)o;
                    }
                }
                else return TFlxFormulaErrorValue.ErrValue;
            }

            return Result;
        }

    }
    #endregion

}
