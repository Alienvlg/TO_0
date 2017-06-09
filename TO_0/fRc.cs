using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace TO_0
{
    public partial class fRc : Form
    {
        public fRc()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
            //this.Close();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {

            string path = @"F:\" + cmbBts.Text + "_" + DateTime.Now.ToString("yyyy-MM-dd");

            try
            {
                if (Directory.Exists(path))
                {
                    MessageBox.Show("Directory already exist");
                    return;
                }

                // Try to create the directory.
                DirectoryInfo dir = Directory.CreateDirectory(path);

                //Delete the directory.
                //dir.Delete(path);
                //Console.WriteLine("The directory was deleted successfully.");
            }
            catch (Exception)
            {
                MessageBox.Show("error1");
            }
            finally { }

            string sourcePath = @"F:\";
            string destinationPath = path;
            string sourceFileName = "Template.docx";
            string destinationFileName = "Учетная карточка.docx";
            string sourceFile = Path.Combine(sourcePath, sourceFileName);
            string destinationFile = Path.Combine(destinationPath, destinationFileName);

            File.Copy(sourceFile, destinationFile, true);

            // Open a WordprocessingDocument for editing using the filepath.
            WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(destinationFile, true);
            
            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
                        
            Paragraph paragraph1 = body.AppendChild(new Paragraph());
            ParagraphProperties Hearder1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "a5" };
            Hearder1.Append(paragraphStyleId1);

            Run run1 = new Run();
            RunProperties runProperties1 = new RunProperties();
            Bold bold2 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = "28" };
            runProperties1.Append(bold2);
            runProperties1.Append(fontSize2);
            Text text1 = new Text();
            text1.Text = "Учетная карточка";
            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(Hearder1);
            paragraph1.Append(run1);

            // add space between paragraph
            Paragraph paragraph2 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            FontSize fontSize3 = new FontSize() { Val = "28" };
            paragraph2.Append(paragraphProperties2);
            //

            Paragraph paragraph3 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "a5" };
            Justification justification1 = new Justification() { Val = JustificationValues.Left };
            Bold bold4 = new Bold();
            FontSize fontSize4 = new FontSize() { Val = "22" };
            Underline underline1 = new Underline() { Val = UnderlineValues.None };
            paragraphProperties3.Append(paragraphStyleId3);
            paragraphProperties3.Append(justification1);


            Run run2 = new Run();
            RunProperties runProperties2 = new RunProperties();
            Bold bold5 = new Bold();
            FontSize fontSize5 = new FontSize() { Val = "22" };
            Underline underline2 = new Underline() { Val = UnderlineValues.None };
            runProperties2.Append(bold5);
            runProperties2.Append(fontSize5);
            runProperties2.Append(underline2);
            Text text2 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text2.Text = "Базовая станция № ";
            run2.Append(runProperties2);
            run2.Append(text2);

            Run run3 = new Run();
            RunProperties runProperties3 = new RunProperties();
            FontSize fontSize6 = new FontSize() { Val = "22" };
            Underline underline3 = new Underline() { Val = UnderlineValues.None };
            runProperties3.Append(fontSize6);
            runProperties3.Append(underline3);
            Text text3 = new Text();
            text3.Text = cmbBts.Text;
            run3.Append(runProperties3);
            run3.Append(text3);

            Run run4 = new Run();
            RunProperties runProperties6 = new RunProperties();
            Bold bold8 = new Bold();
            FontSize fontSize9 = new FontSize() { Val = "22" };
            Underline underline6 = new Underline() { Val = UnderlineValues.None };
            runProperties6.Append(bold8);
            runProperties6.Append(fontSize9);
            runProperties6.Append(underline6);
            Text text6 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text6.Text = ". Название: ";
            run4.Append(runProperties6);
            run4.Append(text6);

            Run run5 = new Run();
            RunProperties runProperties7 = new RunProperties();
            FontSize fontSize10 = new FontSize() { Val = "22" };
            Underline underline7 = new Underline() { Val = UnderlineValues.None };
            runProperties7.Append(fontSize10);
            runProperties7.Append(underline7);
            Text text7 = new Text();
            text7.Text = tbName.Text;
            run5.Append(runProperties7);
            run5.Append(text7);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run2);
            paragraph3.Append(run3);
            paragraph3.Append(run4);
            paragraph3.Append(run5);

            Paragraph paragraph4 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "a5" };
            Justification justification2 = new Justification() { Val = JustificationValues.Left };
            paragraphProperties4.Append(paragraphStyleId4);
            paragraphProperties4.Append(justification2);

            Run run8 = new Run();
            RunProperties runProperties8 = new RunProperties();
            Bold bold9 = new Bold();
            FontSize fontSize12 = new FontSize() { Val = "22" };
            Underline underline9 = new Underline() { Val = UnderlineValues.None };
            runProperties8.Append(bold9);
            runProperties8.Append(fontSize12);
            runProperties8.Append(underline9);
            Text text8 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text8.Text = "Адрес: ";
            run8.Append(runProperties8);
            run8.Append(text8);

            Run run10 = new Run();
            RunProperties runProperties10 = new RunProperties();
            FontSize fontSize14 = new FontSize() { Val = "22" };
            Underline underline11 = new Underline() { Val = UnderlineValues.None };
            runProperties10.Append(fontSize14);
            runProperties10.Append(underline11);
            Text text10 = new Text();
            text10.Text = tbAdress.Text;
            run10.Append(runProperties10);
            run10.Append(text10);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run8);
            paragraph4.Append(run10);

            Paragraph paragraph5 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() { Val = "a5" };
            Justification justification3 = new Justification() { Val = JustificationValues.Left };
            paragraphProperties5.Append(paragraphStyleId5);
            paragraphProperties5.Append(justification3);

            Run run11 = new Run();
            RunProperties runProperties11 = new RunProperties();
            Bold bold10 = new Bold();
            FontSize fontSize16 = new FontSize() { Val = "22" };
            Underline underline13 = new Underline() { Val = UnderlineValues.None };
            runProperties11.Append(bold10);
            runProperties11.Append(fontSize16);
            runProperties11.Append(underline13);
            Text text11 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text11.Text = "Экстренная информация: ";
            run11.Append(runProperties11);
            run11.Append(text11);

            Run run14 = new Run();
            RunProperties runProperties14 = new RunProperties();
            FontSize fontSize19 = new FontSize() { Val = "22" };
            Underline underline16 = new Underline() { Val = UnderlineValues.None };
            runProperties14.Append(fontSize19);
            runProperties14.Append(underline16);
            Text text14 = new Text();
            text14.Text = tbInfo.Text;
            run14.Append(runProperties14);
            run14.Append(text14);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run11);
            paragraph5.Append(run14);

            Paragraph paragraph6 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId6 = new ParagraphStyleId() { Val = "a5" };
            Justification justification4 = new Justification() { Val = JustificationValues.Left };
            paragraphProperties6.Append(paragraphStyleId6);
            paragraphProperties6.Append(justification4);

            Run run15 = new Run();
            RunProperties runProperties15 = new RunProperties();
            Bold bold12 = new Bold();
            FontSize fontSize21 = new FontSize() { Val = "22" };
            Underline underline18 = new Underline() { Val = UnderlineValues.None };
            runProperties15.Append(bold12);
            runProperties15.Append(fontSize21);
            runProperties15.Append(underline18);
            Text text15 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text15.Text = "Место расположения: ";
            run15.Append(runProperties15);
            run15.Append(text15);

            Run run17 = new Run();
            RunProperties runProperties17 = new RunProperties();
            FontSize fontSize23 = new FontSize() { Val = "22" };
            Underline underline20 = new Underline() { Val = UnderlineValues.None };
            runProperties17.Append(fontSize23);
            runProperties17.Append(underline20);
            Text text17 = new Text();
            text17.Text = tbMap.Text;
            run17.Append(runProperties17);
            run17.Append(text17);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run15);
            paragraph6.Append(run17);

            Paragraph paragraph7 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId7 = new ParagraphStyleId() { Val = "a5" };
            Justification justification5 = new Justification() { Val = JustificationValues.Left };
            paragraphProperties7.Append(paragraphStyleId7);
            paragraphProperties7.Append(justification5);

            Run run18 = new Run();
            RunProperties runProperties18 = new RunProperties();
            Bold bold13 = new Bold();
            FontSize fontSize25 = new FontSize() { Val = "22" };
            Underline underline22 = new Underline() { Val = UnderlineValues.None };

            runProperties18.Append(bold13);
            runProperties18.Append(fontSize25);
            runProperties18.Append(underline22);
            Text text18 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text18.Text = "Контактные лица: ";
            run18.Append(runProperties18);
            run18.Append(text18);

            Run run20 = new Run();
            RunProperties runProperties20 = new RunProperties();
            FontSize fontSize27 = new FontSize() { Val = "22" };
            Underline underline24 = new Underline() { Val = UnderlineValues.None };
            runProperties20.Append(fontSize27);
            runProperties20.Append(underline24);
            Text text20 = new Text();
            text20.Text = tbContact.Text;
            run20.Append(runProperties20);
            run20.Append(text20);

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(run18);
            paragraph7.Append(run20);

            Paragraph paragraph8 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId8 = new ParagraphStyleId() { Val = "a5" };
            Justification justification6 = new Justification() { Val = JustificationValues.Left };
            paragraphProperties8.Append(paragraphStyleId8);
            paragraphProperties8.Append(justification6);

            Run run21 = new Run();
            RunProperties runProperties21 = new RunProperties();
            Bold bold14 = new Bold();
            FontSize fontSize29 = new FontSize() { Val = "22" };
            Underline underline26 = new Underline() { Val = UnderlineValues.None };
            runProperties21.Append(bold14);
            runProperties21.Append(fontSize29);
            runProperties21.Append(underline26);
            Text text21 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text21.Text = "Возможность прохода: ";
            run21.Append(runProperties21);
            run21.Append(text21);

            Run run23 = new Run();
            RunProperties runProperties23 = new RunProperties();
            FontSize fontSize31 = new FontSize() { Val = "22" };
            Underline underline28 = new Underline() { Val = UnderlineValues.None };
            runProperties23.Append(fontSize31);
            runProperties23.Append(underline28);
            Text text23 = new Text();
            text23.Text = tbAccess.Text;
            run23.Append(runProperties23);
            run23.Append(text23);

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(run21);
            paragraph8.Append(run23);

            Paragraph paragraph9 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId9 = new ParagraphStyleId() { Val = "a5" };
            Justification justification7 = new Justification() { Val = JustificationValues.Left };
            paragraphProperties9.Append(paragraphStyleId9);
            paragraphProperties9.Append(justification7);

            Run run24 = new Run();
            RunProperties runProperties24 = new RunProperties();
            Bold bold15 = new Bold();
            FontSize fontSize33 = new FontSize() { Val = "22" };
            Underline underline30 = new Underline() { Val = UnderlineValues.None };
            runProperties24.Append(bold15);
            runProperties24.Append(fontSize33);
            runProperties24.Append(underline30);
            Text text24 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text24.Text = "Ограничение по списку: ";
            run24.Append(runProperties24);
            run24.Append(text24);

            Run run27 = new Run();
            RunProperties runProperties27 = new RunProperties();
            FontSize fontSize36 = new FontSize() { Val = "22" };
            Underline underline33 = new Underline() { Val = UnderlineValues.None };
            runProperties27.Append(fontSize36);
            runProperties27.Append(underline33);
            Text text27 = new Text();
            text27.Text = tbAccesList.Text;
            run27.Append(runProperties27);
            run27.Append(text27);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run24);
            paragraph9.Append(run27);

            Paragraph paragraph10 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId10 = new ParagraphStyleId() { Val = "a5" };
            Justification justification8 = new Justification() { Val = JustificationValues.Left };
            FontSize fontSize37 = new FontSize() { Val = "22" };
            Underline underline34 = new Underline() { Val = UnderlineValues.None };
            paragraphProperties10.Append(paragraphStyleId10);
            paragraphProperties10.Append(justification8);

            Run run28 = new Run();
            RunProperties runProperties28 = new RunProperties();
            Bold bold17 = new Bold();
            FontSize fontSize38 = new FontSize() { Val = "22" };
            Underline underline35 = new Underline() { Val = UnderlineValues.None };
            runProperties28.Append(bold17);
            runProperties28.Append(fontSize38);
            runProperties28.Append(underline35);
            Text text28 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text28.Text = "Установка и подключение ДГУ: ";
            run28.Append(runProperties28);
            run28.Append(text28);

            Run run33 = new Run();
            RunProperties runProperties33 = new RunProperties();
            FontSize fontSize43 = new FontSize() { Val = "22" };
            Underline underline40 = new Underline() { Val = UnderlineValues.None };
            runProperties33.Append(fontSize43);
            runProperties33.Append(underline40);
            Text text33 = new Text();
            text33.Text = cmbDgu.Text;
            run33.Append(runProperties33);
            run33.Append(text33);

            paragraph10.Append(paragraphProperties10);
            paragraph10.Append(run28);
            paragraph10.Append(run33);

            Paragraph paragraph11 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId11 = new ParagraphStyleId() { Val = "a5" };
            Justification justification9 = new Justification() { Val = JustificationValues.Left };
            paragraphProperties11.Append(paragraphStyleId11);
            paragraphProperties11.Append(justification9);

            Run run34 = new Run();
            RunProperties runProperties34 = new RunProperties();
            Bold bold21 = new Bold();
            FontSize fontSize45 = new FontSize() { Val = "22" };
            Underline underline42 = new Underline() { Val = UnderlineValues.None };
            runProperties34.Append(bold21);
            runProperties34.Append(fontSize45);
            runProperties34.Append(underline42);
            Text text34 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text34.Text = "Длина кабеля: ";
            run34.Append(runProperties34);
            run34.Append(text34);

            Run run36 = new Run();
            RunProperties runProperties36 = new RunProperties();
            FontSize fontSize47 = new FontSize() { Val = "22" };
            Underline underline44 = new Underline() { Val = UnderlineValues.None };
            runProperties36.Append(fontSize47);
            runProperties36.Append(underline44);
            Text text36 = new Text();
            text36.Text = tbMetr.Text;
            run36.Append(runProperties36);
            run36.Append(text36);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run34);
            paragraph11.Append(run36);

            Paragraph paragraph12 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId12 = new ParagraphStyleId() { Val = "a5" };
            Justification justification10 = new Justification() { Val = JustificationValues.Left };
            paragraphProperties12.Append(paragraphStyleId12);
            paragraphProperties12.Append(justification10);

            Run run37 = new Run();
            RunProperties runProperties37 = new RunProperties();
            Bold bold22 = new Bold();
            FontSize fontSize49 = new FontSize() { Val = "22" };
            Underline underline46 = new Underline() { Val = UnderlineValues.None };
            runProperties37.Append(bold22);
            runProperties37.Append(fontSize49);
            runProperties37.Append(underline46);
            Text text37 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text37.Text = "Электроснабжение осуществляется от: ";
            run37.Append(runProperties37);
            run37.Append(text37);

            Run run39 = new Run();
            RunProperties runProperties39 = new RunProperties();
            FontSize fontSize51 = new FontSize() { Val = "22" };
            Underline underline48 = new Underline() { Val = UnderlineValues.None };
            runProperties39.Append(fontSize51);
            runProperties39.Append(underline48);
            Text text39 = new Text();
            text39.Text = tbPwr.Text;
            run39.Append(runProperties39);
            run39.Append(text39);

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(run37);
            paragraph12.Append(run39);

            Paragraph paragraph13 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId13 = new ParagraphStyleId() { Val = "a5" };
            Justification justification11 = new Justification() { Val = JustificationValues.Left };
            paragraphProperties13.Append(paragraphStyleId13);
            paragraphProperties13.Append(justification11);

            Run run40 = new Run();
            RunProperties runProperties40 = new RunProperties();
            Bold bold23 = new Bold();
            FontSize fontSize53 = new FontSize() { Val = "22" };
            Underline underline50 = new Underline() { Val = UnderlineValues.None };
            runProperties40.Append(bold23);
            runProperties40.Append(fontSize53);
            runProperties40.Append(underline50);
            Text text40 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text40.Text = "Ключ от аппаратной: ";
            run40.Append(runProperties40);
            run40.Append(text40);
            
            Run run42 = new Run();
            RunProperties runProperties42 = new RunProperties();
            FontSize fontSize55 = new FontSize() { Val = "22" };
            Underline underline52 = new Underline() { Val = UnderlineValues.None };
            runProperties42.Append(fontSize55);
            runProperties42.Append(underline52);
            Text text42 = new Text();
            text42.Text = tbKey.Text;
            run42.Append(runProperties42);
            run42.Append(text42);

            paragraph13.Append(paragraphProperties13);
            paragraph13.Append(run40);
            paragraph13.Append(run42);

            Paragraph paragraph14 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId14 = new ParagraphStyleId() { Val = "a5" };
            Justification justification12 = new Justification() { Val = JustificationValues.Left };
            paragraphProperties14.Append(paragraphStyleId14);
            paragraphProperties14.Append(justification12);

            Run run43 = new Run();
            RunProperties runProperties43 = new RunProperties();
            Bold bold24 = new Bold();
            FontSize fontSize57 = new FontSize() { Val = "22" };
            Underline underline54 = new Underline() { Val = UnderlineValues.None };
            runProperties43.Append(bold24);
            runProperties43.Append(fontSize57);
            runProperties43.Append(underline54);
            Text text43 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text43.Text = "Выход на крышу: ";
            run43.Append(runProperties43);
            run43.Append(text43);

            Run run47 = new Run();
            RunProperties runProperties47 = new RunProperties();
            FontSize fontSize61 = new FontSize() { Val = "22" };
            Underline underline58 = new Underline() { Val = UnderlineValues.None };
            runProperties47.Append(fontSize61);
            runProperties47.Append(underline58);
            Text text47 = new Text();
            text47.Text = tbRoof.Text;
            run47.Append(runProperties47);
            run47.Append(text47);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(run43);
            paragraph14.Append(run47);

            // add space between paragraph
            Paragraph paragraph15 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties15 = new ParagraphProperties();
            FontSize fontSize7 = new FontSize() { Val = "28" };
            paragraph15.Append(paragraphProperties15);
            //

            Paragraph paragraph16 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId16 = new ParagraphStyleId() { Val = "a5" };
            paragraphProperties16.Append(paragraphStyleId16);

            Run run48 = new Run();
            RunProperties runProperties48 = new RunProperties();
            Bold bold28 = new Bold();
            FontSize fontSize63 = new FontSize() { Val = "28" };
            runProperties48.Append(bold28);
            runProperties48.Append(fontSize63);
            Text text48 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text48.Text = "Оборудование BTS";
            run48.Append(runProperties48);
            run48.Append(text48);
            paragraph16.Append(paragraphProperties16);
            paragraph16.Append(run48);

            // add space between paragraph
            Paragraph paragraph17 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties17 = new ParagraphProperties();
            FontSize fontSize64 = new FontSize() { Val = "28" };
            paragraph17.Append(paragraphProperties17);
            //

            Table table1 = new Table();
            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "9360", Type = TableWidthUnitValues.Dxa };
            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() {Val = BorderValues.Single, Size = 10};
            LeftBorder leftBorder1 = new LeftBorder() {Val = BorderValues.Single, Size = 10 };
            BottomBorder bottomBorder1 = new BottomBorder() {Val = BorderValues.Single, Size = 10 };
            RightBorder rightBorder1 = new RightBorder() {Val = BorderValues.Single, Size = 10 };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() {Val = BorderValues.Single, Size = 10 };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() {Val = BorderValues.Single, Size = 10 };
            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            tableProperties1.Append(tableBorders1);

            TableRow tableRow1 = new TableRow();

            TableCell tableCell1 = new TableCell();
            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "2250", Type = TableWidthUnitValues.Dxa };
            tableCellProperties1.Append(tableCellWidth1);
            Paragraph paragraph18 = new Paragraph();
            ParagraphProperties paragraphProperties18 = new ParagraphProperties();
            Justification justification13 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties18.Append(justification13);
            Run run51 = new Run() { RsidRunProperties = "0017479D" };
            RunProperties runProperties51 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "Calibri" };
            FontSize fontSize68 = new FontSize() { Val = "18" };
            runProperties51.Append(runFonts4);
            runProperties51.Append(fontSize68);
            Text text51 = new Text();
            text51.Text = "Тип шкафа";
            run51.Append(runProperties51);
            run51.Append(text51);
            paragraph18.Append(paragraphProperties18);
            paragraph18.Append(run51);
            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph18);


            TableCell tableCell2 = new TableCell();
            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "3690", Type = TableWidthUnitValues.Dxa };
            tableCellProperties2.Append(tableCellWidth2);
            Paragraph paragraph19 = new Paragraph();
            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            Justification justification14 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties19.Append(justification14);
            Run run52 = new Run() { RsidRunProperties = "0017479D" };
            RunProperties runProperties52 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "Calibri" };
            FontSize fontSize70 = new FontSize() { Val = "18" };
            runProperties52.Append(runFonts6);
            runProperties52.Append(fontSize70);
            Text text52 = new Text();
            text52.Text = "Конфигурация";
            run52.Append(runProperties52);
            run52.Append(text52);
            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run52);
            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph19);

            TableCell tableCell3 = new TableCell();
            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "1800", Type = TableWidthUnitValues.Dxa };
            tableCellProperties3.Append(tableCellWidth3);
            Paragraph paragraph20 = new Paragraph();
            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            Justification justification15 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties20.Append(justification15);
            Run run53 = new Run();
            RunProperties runProperties53 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "Calibri" };
            FontSize fontSize72 = new FontSize() { Val = "18" };
            runProperties53.Append(runFonts8);
            runProperties53.Append(fontSize72);
            Text text53 = new Text();
            text53.Text = "S/N";
            run53.Append(runProperties53);
            run53.Append(text53);
            paragraph20.Append(paragraphProperties20);
            paragraph20.Append(run53);
            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph20);

            TableCell tableCell4 = new TableCell();
            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "1620", Type = TableWidthUnitValues.Dxa };
            tableCellProperties4.Append(tableCellWidth4);
            Paragraph paragraph21 = new Paragraph();
            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            Justification justification16 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties21.Append(justification16);
            Run run54 = new Run() { RsidRunProperties = "0017479D" };
            RunProperties runProperties54 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "Calibri" };
            FontSize fontSize74 = new FontSize() { Val = "18" };
            runProperties54.Append(runFonts10);
            runProperties54.Append(fontSize74);
            Text text54 = new Text();
            text54.Text = "Дата установки";
            run54.Append(runProperties54);
            run54.Append(text54);
            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(run54);
            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph21);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);
            tableRow1.Append(tableCell4);
            table1.Append(tableProperties1);
            table1.Append(tableRow1);

            wordprocessingDocument.MainDocumentPart.Document.Body.Append(table1);

            // add space between paragraph
            Paragraph paragraph22 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            FontSize fontSize67 = new FontSize() { Val = "28" };
            paragraph17.Append(paragraphProperties22);
            //

            Paragraph paragraph39 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties39 = new ParagraphProperties();
            Justification justification29 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties39.Append(justification29);
            Run run60 = new Run();
            RunProperties runProperties60 = new RunProperties();
            Bold bold43 = new Bold();
            FontSize fontSize98 = new FontSize() { Val = "28" };
            Underline underline62 = new Underline() { Val = UnderlineValues.Single };
            runProperties60.Append(bold43);
            runProperties60.Append(fontSize98);
            runProperties60.Append(underline62);
            Text text60 = new Text();
            text60.Text = "Система электропитания";
            run60.Append(runProperties60);
            run60.Append(text60);
            paragraph39.Append(paragraphProperties39);
            paragraph39.Append(run60);

            // add space between paragraph
            Paragraph paragraph40 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties40 = new ParagraphProperties();
            FontSize fontSize99 = new FontSize() { Val = "28" };
            paragraph40.Append(paragraphProperties40);
            //

            Table table2 = new Table();
            TableProperties tableProperties2 = new TableProperties();
            TableWidth tableWidth2 = new TableWidth() { Width = "9560", Type = TableWidthUnitValues.Dxa };
            TableBorders tableBorders2 = new TableBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Size = 10 };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Size = 10 };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Size = 10 };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Size = 10 };
            InsideHorizontalBorder insideHorizontalBorder2 = new InsideHorizontalBorder() { Val = BorderValues.Single, Size = 10 };
            InsideVerticalBorder insideVerticalBorder2 = new InsideVerticalBorder() { Val = BorderValues.Single, Size = 10 };
            tableBorders2.Append(topBorder2);
            tableBorders2.Append(leftBorder2);
            tableBorders2.Append(bottomBorder2);
            tableBorders2.Append(rightBorder2);
            tableBorders2.Append(insideHorizontalBorder2);
            tableBorders2.Append(insideVerticalBorder2);
            
            tableProperties2.Append(tableWidth2);
            tableProperties2.Append(tableBorders2);

            TableRow tableRow6 = new TableRow();
            TableCell tableCell21 = new TableCell();
            TableCellProperties tableCellProperties21 = new TableCellProperties();
            TableCellWidth tableCellWidth21 = new TableCellWidth() { Width = "2875", Type = TableWidthUnitValues.Dxa };
            tableCellProperties21.Append(tableCellWidth21);
            Paragraph paragraph41 = new Paragraph();
            ParagraphProperties paragraphProperties41 = new ParagraphProperties();
            Justification justification30 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties41.Append(justification30);
            Run run61 = new Run();
            RunProperties runProperties61 = new RunProperties();
            RunFonts runFonts34 = new RunFonts() { Ascii = "Calibri" };
            FontSize fontSize101 = new FontSize() { Val = "18" };
            runProperties61.Append(runFonts34);
            runProperties61.Append(fontSize101);
            Text text61 = new Text();
            text61.Text = "Тип шкафа";
            run61.Append(runProperties61);
            run61.Append(text61);
            paragraph41.Append(paragraphProperties41);
            paragraph41.Append(run61);
            tableCell21.Append(tableCellProperties21);
            tableCell21.Append(paragraph41);

            TableCell tableCell22 = new TableCell();
            TableCellProperties tableCellProperties22 = new TableCellProperties();
            TableCellWidth tableCellWidth22 = new TableCellWidth() { Width = "3245", Type = TableWidthUnitValues.Dxa };
            tableCellProperties22.Append(tableCellWidth22);
            Paragraph paragraph42 = new Paragraph();
            ParagraphProperties paragraphProperties42 = new ParagraphProperties();
            Justification justification31 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties42.Append(justification31);
            Run run62 = new Run();
            RunProperties runProperties62 = new RunProperties();
            RunFonts runFonts36 = new RunFonts() { Ascii = "Calibri" };
            FontSize fontSize103 = new FontSize() { Val = "18" };
            runProperties62.Append(runFonts36);
            runProperties62.Append(fontSize103);
            Text text62 = new Text();
            text62.Text = "Блок управления";
            run62.Append(runProperties62);
            run62.Append(text62);
            paragraph42.Append(paragraphProperties42);
            paragraph42.Append(run62);
            tableCell22.Append(tableCellProperties22);
            tableCell22.Append(paragraph42);

            TableCell tableCell23 = new TableCell();
            TableCellProperties tableCellProperties23 = new TableCellProperties();
            TableCellWidth tableCellWidth23 = new TableCellWidth() { Width = "1620", Type = TableWidthUnitValues.Dxa };
            tableCellProperties23.Append(tableCellWidth23);
            Paragraph paragraph43 = new Paragraph();
            ParagraphProperties paragraphProperties43 = new ParagraphProperties();
            Justification justification32 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties43.Append(justification32);
            Run run73 = new Run();
            RunProperties runProperties73 = new RunProperties();
            RunFonts runFonts48 = new RunFonts() { Ascii = "Calibri" };
            FontSize fontSize115 = new FontSize() { Val = "18" };
            runProperties73.Append(runFonts48);
            runProperties73.Append(fontSize115);
            Text text73 = new Text();
            text73.Text = "Напряжение";
            run73.Append(runProperties73);
            run73.Append(text73);
            paragraph43.Append(paragraphProperties43);
            paragraph43.Append(run73);
            tableCell23.Append(tableCellProperties23);
            tableCell23.Append(paragraph43);

            TableCell tableCell24 = new TableCell();
            TableCellProperties tableCellProperties24 = new TableCellProperties();
            TableCellWidth tableCellWidth24 = new TableCellWidth() { Width = "1820", Type = TableWidthUnitValues.Dxa };
            tableCellProperties24.Append(tableCellWidth24);
            Paragraph paragraph44 = new Paragraph();
            ParagraphProperties paragraphProperties44 = new ParagraphProperties();
            Justification justification33 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties44.Append(justification33);
            Run run74 = new Run();
            RunProperties runProperties74 = new RunProperties();
            RunFonts runFonts50 = new RunFonts() { Ascii = "Calibri" };
            FontSize fontSize117 = new FontSize() { Val = "18" };
            runProperties74.Append(runFonts50);
            runProperties74.Append(fontSize117);
            Text text74 = new Text();
            text74.Text = "S/N шкафа";
            run74.Append(runProperties74);
            run74.Append(text74);
            paragraph44.Append(paragraphProperties44);
            paragraph44.Append(run74);
            tableCell24.Append(tableCellProperties24);
            tableCell24.Append(paragraph44);

            tableRow6.Append(tableCell21);
            tableRow6.Append(tableCell22);
            tableRow6.Append(tableCell23);
            tableRow6.Append(tableCell24);

            TableRow tableRow7 = new TableRow();
            TableCell tableCell25 = new TableCell();
            TableCellProperties tableCellProperties25 = new TableCellProperties();
            TableCellWidth tableCellWidth25 = new TableCellWidth() { Width = "2875", Type = TableWidthUnitValues.Dxa };
            tableCellProperties25.Append(tableCellWidth25);
            Paragraph paragraph45 = new Paragraph();
            ParagraphProperties paragraphProperties45 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines25 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.AtLeast };
            Indentation indentation18 = new Indentation() { Left = "15" };
            paragraphProperties45.Append(spacingBetweenLines25);
            paragraphProperties45.Append(indentation18);

            Run run75 = new Run();
            RunProperties runProperties75 = new RunProperties();
            RunFonts runFonts52 = new RunFonts() { Ascii = "Calibri" };
            FontSize fontSize119 = new FontSize() { Val = "18" };
            runProperties75.Append(runFonts52);
            runProperties75.Append(fontSize119);
            Text text75 = new Text();
            text75.Text = "Delta Energy System";
            run75.Append(runProperties75);
            run75.Append(text75);
            paragraph45.Append(paragraphProperties45);
            paragraph45.Append(run75);
            tableCell25.Append(tableCellProperties25);
            tableCell25.Append(paragraph45);

            TableCell tableCell26 = new TableCell();
            TableCellProperties tableCellProperties26 = new TableCellProperties();
            TableCellWidth tableCellWidth26 = new TableCellWidth() { Width = "3245", Type = TableWidthUnitValues.Dxa };
            tableCellProperties26.Append(tableCellWidth26);
            Paragraph paragraph46 = new Paragraph();
            ParagraphProperties paragraphProperties46 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines26 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.AtLeast };
            Justification justification34 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties46.Append(spacingBetweenLines26);
            paragraphProperties46.Append(justification34);
            paragraph46.Append(paragraphProperties46);
            tableCell26.Append(tableCellProperties26);
            tableCell26.Append(paragraph46);

            TableCell tableCell27 = new TableCell();
            TableCellProperties tableCellProperties27 = new TableCellProperties();
            TableCellWidth tableCellWidth27 = new TableCellWidth() { Width = "1620", Type = TableWidthUnitValues.Dxa };
            tableCellProperties27.Append(tableCellWidth27);
            Paragraph paragraph47 = new Paragraph();
            ParagraphProperties paragraphProperties47 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines27 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.AtLeast };
            Justification justification35 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties47.Append(spacingBetweenLines27);
            paragraphProperties47.Append(justification35);
            Run run76 = new Run();
            RunProperties runProperties76 = new RunProperties();
            RunFonts runFonts55 = new RunFonts() { Ascii = "Calibri" };
            FontSize fontSize122 = new FontSize() { Val = "18" };
            runProperties76.Append(runFonts55);
            runProperties76.Append(fontSize122);
            Text text76 = new Text();
            text76.Text = "54.5";
            run76.Append(runProperties76);
            run76.Append(text76);
            paragraph47.Append(paragraphProperties47);
            paragraph47.Append(run76);
            tableCell27.Append(tableCellProperties27);
            tableCell27.Append(paragraph47);

            TableCell tableCell28 = new TableCell();
            TableCellProperties tableCellProperties28 = new TableCellProperties();
            TableCellWidth tableCellWidth28 = new TableCellWidth() { Width = "1820", Type = TableWidthUnitValues.Dxa };
            tableCellProperties28.Append(tableCellWidth28);
            Paragraph paragraph48 = new Paragraph();
            ParagraphProperties paragraphProperties48 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines28 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.AtLeast };
            Justification justification36 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties48.Append(spacingBetweenLines28);
            paragraphProperties48.Append(justification36);
            Run run77 = new Run();
            RunProperties runProperties77 = new RunProperties();
            RunFonts runFonts57 = new RunFonts() { Ascii = "Calibri" };
            FontSize fontSize124 = new FontSize() { Val = "18" };
            runProperties77.Append(runFonts57);
            runProperties77.Append(fontSize124);
            Text text77 = new Text();
            text77.Text = "5001155";
            run77.Append(runProperties77);
            run77.Append(text77);
            paragraph48.Append(paragraphProperties48);
            paragraph48.Append(run77);
            tableCell28.Append(tableCellProperties28);
            tableCell28.Append(paragraph48);

            tableRow7.Append(tableCell25);
            tableRow7.Append(tableCell26);
            tableRow7.Append(tableCell27);
            tableRow7.Append(tableCell28);

            TableRow tableRow8 = new TableRow();
            TableCell tableCell29 = new TableCell();
            TableCellProperties tableCellProperties29 = new TableCellProperties();
            TableCellWidth tableCellWidth29 = new TableCellWidth() { Width = "2875", Type = TableWidthUnitValues.Dxa };
            tableCellProperties29.Append(tableCellWidth29);
            Paragraph paragraph49 = new Paragraph();
            ParagraphProperties paragraphProperties49 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines29 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.AtLeast };
            Justification justification37 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties49.Append(spacingBetweenLines29);
            paragraphProperties49.Append(justification37);
            paragraph49.Append(paragraphProperties49);
            tableCell29.Append(tableCellProperties29);
            tableCell29.Append(paragraph49);

            TableCell tableCell30 = new TableCell();
            TableCellProperties tableCellProperties30 = new TableCellProperties();
            TableCellWidth tableCellWidth30 = new TableCellWidth() { Width = "3245", Type = TableWidthUnitValues.Dxa };
            tableCellProperties30.Append(tableCellWidth30);
            Paragraph paragraph50 = new Paragraph();
            ParagraphProperties paragraphProperties50 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines30 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.AtLeast };
            Justification justification38 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties50.Append(spacingBetweenLines30);
            paragraphProperties50.Append(justification38);
            paragraph50.Append(paragraphProperties50);
            tableCell30.Append(tableCellProperties30);
            tableCell30.Append(paragraph50);

            TableCell tableCell31 = new TableCell();
            TableCellProperties tableCellProperties31 = new TableCellProperties();
            TableCellWidth tableCellWidth31 = new TableCellWidth() { Width = "1620", Type = TableWidthUnitValues.Dxa };
            tableCellProperties31.Append(tableCellWidth31);
            Paragraph paragraph51 = new Paragraph();
            ParagraphProperties paragraphProperties51 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines31 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.AtLeast };
            Justification justification39 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties51.Append(spacingBetweenLines31);
            paragraphProperties51.Append(justification39);
            paragraph51.Append(paragraphProperties51);
            tableCell31.Append(tableCellProperties31);
            tableCell31.Append(paragraph51);

            TableCell tableCell32 = new TableCell();
            TableCellProperties tableCellProperties32 = new TableCellProperties();
            TableCellWidth tableCellWidth32 = new TableCellWidth() { Width = "1820", Type = TableWidthUnitValues.Dxa };
            tableCellProperties32.Append(tableCellWidth32);
            Paragraph paragraph52 = new Paragraph();
            ParagraphProperties paragraphProperties52 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines32 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.AtLeast };
            Justification justification40 = new Justification() { Val = JustificationValues.Center };
            paragraphProperties52.Append(spacingBetweenLines32);
            paragraphProperties52.Append(justification40);
            paragraph52.Append(paragraphProperties52);
            tableCell32.Append(tableCellProperties32);
            tableCell32.Append(paragraph52);

            tableRow8.Append(tableCell29);
            tableRow8.Append(tableCell30);
            tableRow8.Append(tableCell31);
            tableRow8.Append(tableCell32);
            table2.Append(tableProperties2);
            table2.Append(tableRow6);
            table2.Append(tableRow7);
            table2.Append(tableRow8);
            wordprocessingDocument.MainDocumentPart.Document.Body.Append(table2);

            // add space between paragraph
            Paragraph paragraph53 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties53 = new ParagraphProperties();
            FontSize fontSize129 = new FontSize() { Val = "28" };
            paragraph53.Append(paragraphProperties53);
            //

            Paragraph paragraph54 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties54 = new ParagraphProperties();
            Justification justification154 = new Justification() { Val = JustificationValues.Left };
            FontSize fontSize54 = new FontSize() { Val = "22" };
            paragraphProperties54.Append(justification154);

            Run run78 = new Run();
            RunProperties runProperties78 = new RunProperties();
            FontSize fontSize131 = new FontSize() { Val = "22" };
            runProperties78.Append(fontSize131);
            Text text78 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text78.Text = "Тип вводного эл. щита: ";
            run78.Append(runProperties78);
            run78.Append(text78);

            Run run83 = new Run();
            RunProperties runProperties83 = new RunProperties();
            FontSize fontSize137 = new FontSize() { Val = "22" };
            runProperties83.Append(fontSize137);
            Text text82 = new Text();
            text82.Text = "----------------";
            run83.Append(runProperties83);
            run83.Append(text82);


            Run run79 = new Run();
            RunProperties runProperties79 = new RunProperties();
            FontSize fontSize132 = new FontSize() { Val = "22" };
            runProperties79.Append(fontSize132);
            TabChar tabChar1 = new TabChar();
            run79.Append(runProperties79);
            run79.Append(tabChar1);

            Run run80 = new Run();
            RunProperties runProperties80 = new RunProperties();
            FontSize fontSize133 = new FontSize() { Val = "22" };
            runProperties80.Append(fontSize133);
            Text text79 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text79.Text = "Розетка для бензогенератора: ";
            run80.Append(runProperties80);
            run80.Append(text79);

            Run run81 = new Run();
            RunProperties runProperties81 = new RunProperties();
            FontSize fontSize134 = new FontSize() { Val = "22" };
            runProperties81.Append(fontSize134);
            Text text80 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text80.Text = "----------";
            run81.Append(runProperties81);
            run81.Append(text80);
            paragraph54.Append(paragraphProperties54);
            paragraph54.Append(run78);
            paragraph54.Append(run79);
            paragraph54.Append(run80);
            paragraph54.Append(run81);
            paragraph54.Append(run83);

            Paragraph paragraph55 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties55 = new ParagraphProperties();
            //FontSize fontSize135 = new FontSize() { Val = "22" };

            Run run82 = new Run();
            RunProperties runProperties82 = new RunProperties();
            FontSize fontSize136 = new FontSize() { Val = "22" };
            runProperties82.Append(fontSize136);
            Text text81 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text81.Text = "Количество выпрямительных модулей ЭПУ 1: ";
            run82.Append(runProperties82);
            run82.Append(text81);

            Run run87 = new Run() { RsidRunAddition = "0017479D" };
            RunProperties runProperties87 = new RunProperties();
            FontSize fontSize141 = new FontSize() { Val = "22" };
            runProperties87.Append(fontSize141);
            Text text86 = new Text();
            text86.Text = "----";
            run87.Append(runProperties87);
            run87.Append(text86);

            Run run88 = new Run();
            RunProperties runProperties88 = new RunProperties();
            FontSize fontSize142 = new FontSize() { Val = "22" };
            runProperties88.Append(fontSize142);
            TabChar tabChar2 = new TabChar();
            run88.Append(runProperties88);
            run88.Append(tabChar2);

            Run run91 = new Run();
            RunProperties runProperties91 = new RunProperties();
            FontSize fontSize145 = new FontSize() { Val = "22" };
            runProperties91.Append(fontSize145);
            Text text89 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text89.Text = "; мощность: ";
            run91.Append(runProperties91);
            run91.Append(text89);

            Run run92 = new Run();
            RunProperties runProperties92 = new RunProperties();
            FontSize fontSize146 = new FontSize() { Val = "22" };
            runProperties92.Append(fontSize146);
            Text text90 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text90.Text = "----";
            run92.Append(runProperties92);
            run92.Append(text90);

            Run run93 = new Run();
            RunProperties runProperties93 = new RunProperties();
            FontSize fontSize147 = new FontSize() { Val = "22" };
            runProperties93.Append(fontSize147);
            Text text91 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text91.Text = " Ватт";
            run93.Append(runProperties93);
            run93.Append(text91);

            paragraph55.Append(paragraphProperties55);
            paragraph55.Append(run82);
            paragraph55.Append(run87);
            paragraph55.Append(run88);
            paragraph55.Append(run91);
            paragraph55.Append(run92);
            paragraph55.Append(run93);


            Paragraph paragraph56 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties56 = new ParagraphProperties();
            FontSize fontSize148 = new FontSize() { Val = "22" };
            Run run94 = new Run();
            RunProperties runProperties94 = new RunProperties();
            FontSize fontSize149 = new FontSize() { Val = "22" };
            runProperties94.Append(fontSize149);
            Text text92 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text92.Text = "Количество выпрямительных модулей ЭПУ 2: ";
            run94.Append(runProperties94);
            run94.Append(text92);

            Run run95 = new Run();
            RunProperties runProperties95 = new RunProperties();
            FontSize fontSize150 = new FontSize() { Val = "22" };
            runProperties95.Append(fontSize150);
            Text text93 = new Text();
            text93.Text = "----";
            run95.Append(runProperties95);
            run95.Append(text93);

            Run run99 = new Run();
            RunProperties runProperties99 = new RunProperties();
            FontSize fontSize154 = new FontSize() { Val = "22" };
            runProperties99.Append(fontSize154);
            TabChar tabChar3 = new TabChar();
            run99.Append(runProperties99);
            run99.Append(tabChar3);

            Run run96 = new Run();
            RunProperties runProperties96 = new RunProperties();
            FontSize fontSize151 = new FontSize() { Val = "22" };
            runProperties96.Append(fontSize151);
            Text text94 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text94.Text = "; мощность: ";
            run96.Append(runProperties96);
            run96.Append(text94);

            Run run98 = new Run();
            RunProperties runProperties98 = new RunProperties();
            FontSize fontSize153 = new FontSize() { Val = "22" };
            runProperties98.Append(fontSize153);
            Text text96 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text96.Text = "----";
            run98.Append(runProperties98);
            run98.Append(text96);

            Run run102 = new Run();
            RunProperties runProperties102 = new RunProperties();
            FontSize fontSize157 = new FontSize() { Val = "22" };
            runProperties102.Append(fontSize157);
            Text text99 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text99.Text = " Ватт";
            run102.Append(runProperties102);
            run102.Append(text99);

            paragraph56.Append(paragraphProperties56);
            paragraph56.Append(run94);
            paragraph56.Append(run95);
            paragraph56.Append(run99);
            paragraph56.Append(run96);
            paragraph56.Append(run98);
            paragraph56.Append(run102);

            Paragraph paragraph57 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties57 = new ParagraphProperties();

            Run run103 = new Run();
            RunProperties runProperties103 = new RunProperties();
            FontSize fontSize159 = new FontSize() { Val = "22" };
            runProperties103.Append(fontSize159);
            Text text100 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text100.Text = "Тип аккумуляторных батарей ЭПУ 1: ";
            run103.Append(runProperties103);
            run103.Append(text100);

            Run run108 = new Run();
            RunProperties runProperties108 = new RunProperties();
            FontSize fontSize164 = new FontSize() { Val = "22" };
            runProperties108.Append(fontSize164);
            Text text105 = new Text();
            text105.Text = "-------";
            run108.Append(runProperties108);
            run108.Append(text105);
            paragraph57.Append(paragraphProperties57);
            paragraph57.Append(run103);
            paragraph57.Append(run108);

            Paragraph paragraph58 = body.AppendChild(new Paragraph());
            ParagraphProperties paragraphProperties58 = new ParagraphProperties();

            Run run109 = new Run();
            RunProperties runProperties109 = new RunProperties();
            FontSize fontSize160 = new FontSize() { Val = "22" };
            runProperties109.Append(fontSize160);
            Text text106 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text106.Text = "Тип аккумуляторных батарей ЭПУ 2: ";
            run109.Append(runProperties109);
            run109.Append(text106);

            Run run112 = new Run();
            RunProperties runProperties112 = new RunProperties();
            FontSize fontSize161 = new FontSize() { Val = "22" };
            runProperties112.Append(fontSize161);
            Text text109 = new Text();
            text109.Text = "-------";
            run112.Append(runProperties112);
            run112.Append(text109);
            paragraph58.Append(paragraphProperties58);
            paragraph58.Append(run109);
            paragraph58.Append(run112);

            Paragraph paragraph59 = body.AppendChild(new Paragraph());
            //ParagraphProperties paragraphProperties59 = new ParagraphProperties();

            Run run113 = new Run();
            RunProperties runProperties113 = new RunProperties();
            FontSize fontSize162 = new FontSize() { Val = "22" };
            runProperties113.Append(fontSize162);
            Text text110 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text110.Text = "Суммарная емкость аккумуляторных батарей ЭПУ 1: ";
            run113.Append(runProperties113);
            run113.Append(text110);

            Run run118 = new Run();
            RunProperties runProperties114 = new RunProperties();
            FontSize fontSize163 = new FontSize() { Val = "22" };
            runProperties114.Append(fontSize163);
            //RunProperties runProperties118 = new RunProperties();
            Text text115 = new Text();
            text115.Text = "-----";
            run118.Append(runProperties114);
            run118.Append(text115);
            //paragraph59.Append(paragraphProperties59);
            paragraph59.Append(run113);
            paragraph59.Append(run118);

            Paragraph paragraph60 = body.AppendChild(new Paragraph());

            Run run119 = new Run();
            RunProperties runProperties119 = new RunProperties();
            FontSize fontSize165 = new FontSize() { Val = "22" };
            runProperties119.Append(fontSize165);
            Text text116 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text116.Text = "Суммарная емкость аккумуляторных батарей ЭПУ 2: ";
            run119.Append(runProperties119);
            run119.Append(text116);

            Run run122 = new Run();
            RunProperties runProperties122 = new RunProperties();
            Text text119 = new Text();
            text119.Text = "-----";
            run122.Append(runProperties122);
            run122.Append(text119);

            paragraph60.Append(run119);
            paragraph60.Append(run122);


            ///////////////////////////////////////////////////////////

            Paragraph paragraph61 = new Paragraph() { RsidParagraphMarkRevision = "0017479D", RsidParagraphAddition = "002031B2", RsidParagraphProperties = "002031B2", RsidRunAdditionDefault = "002031B2" };

            ParagraphProperties paragraphProperties61 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId21 = new ParagraphStyleId() { Val = "1" };

            ParagraphMarkRunProperties paragraphMarkRunProperties60 = new ParagraphMarkRunProperties();
            Bold bold59 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript65 = new BoldComplexScript() { Val = false };
            //Color color18 = new Color() { Val = "auto" };
            FontSizeComplexScript fontSizeComplexScript176 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties60.Append(bold59);
            paragraphMarkRunProperties60.Append(boldComplexScript65);
           // paragraphMarkRunProperties60.Append(color18);
            paragraphMarkRunProperties60.Append(fontSizeComplexScript176);

            paragraphProperties61.Append(paragraphStyleId21);
            paragraphProperties61.Append(paragraphMarkRunProperties60);

            Run run123 = new Run() { RsidRunProperties = "0017479D" };

            RunProperties runProperties123 = new RunProperties();
            Bold bold60 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript66 = new BoldComplexScript() { Val = false };
            //Color color19 = new Color() { Val = "auto" };
            FontSizeComplexScript fontSizeComplexScript177 = new FontSizeComplexScript() { Val = "22" };

            runProperties123.Append(bold60);
            runProperties123.Append(boldComplexScript66);
            //runProperties123.Append(color19);
            runProperties123.Append(fontSizeComplexScript177);
            Text text120 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text120.Text = "Ток потребления ";

            run123.Append(runProperties123);
            run123.Append(text120);

            Run run124 = new Run() { RsidRunProperties = "0017479D", RsidRunAddition = "004E6FEF" };

            RunProperties runProperties124 = new RunProperties();
            Bold bold61 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript67 = new BoldComplexScript() { Val = false };
            //Color color20 = new Color() { Val = "auto" };
            FontSizeComplexScript fontSizeComplexScript178 = new FontSizeComplexScript() { Val = "22" };

            runProperties124.Append(bold61);
            runProperties124.Append(boldComplexScript67);
            //runProperties124.Append(color20);
            runProperties124.Append(fontSizeComplexScript178);
            Text text121 = new Text();
            text121.Text = "ЭПУ";

            run124.Append(runProperties124);
            run124.Append(text121);

            Run run125 = new Run() { RsidRunAddition = "0017479D" };

            RunProperties runProperties125 = new RunProperties();
            Bold bold62 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript68 = new BoldComplexScript() { Val = false };
            //Color color21 = new Color() { Val = "auto" };
            FontSizeComplexScript fontSizeComplexScript179 = new FontSizeComplexScript() { Val = "22" };

            runProperties125.Append(bold62);
            runProperties125.Append(boldComplexScript68);
            //runProperties125.Append(color21);
            runProperties125.Append(fontSizeComplexScript179);
            Text text122 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text122.Text = " 1:-----";

            run125.Append(runProperties125);
            run125.Append(text122);

            Run run126 = new Run() { RsidRunProperties = "003E08A7", RsidRunAddition = "0017479D" };

            RunProperties runProperties126 = new RunProperties();
            Bold bold63 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript69 = new BoldComplexScript() { Val = false };
            //Color color22 = new Color() { Val = "auto" };
            FontSizeComplexScript fontSizeComplexScript180 = new FontSizeComplexScript() { Val = "22" };

            runProperties126.Append(bold63);
            runProperties126.Append(boldComplexScript69);
            //runProperties126.Append(color22);
            runProperties126.Append(fontSizeComplexScript180);
            TabChar tabChar4 = new TabChar();

            run126.Append(runProperties126);
            run126.Append(tabChar4);

            Run run127 = new Run() { RsidRunProperties = "0017479D" };

            RunProperties runProperties127 = new RunProperties();
            Bold bold64 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript70 = new BoldComplexScript() { Val = false };
            //Color color23 = new Color() { Val = "auto" };
            FontSizeComplexScript fontSizeComplexScript181 = new FontSizeComplexScript() { Val = "22" };
            Languages languages46 = new Languages() { Val = "en-US" };

            runProperties127.Append(bold64);
            runProperties127.Append(boldComplexScript70);
            //runProperties127.Append(color23);
            runProperties127.Append(fontSizeComplexScript181);
            runProperties127.Append(languages46);
            Text text123 = new Text();
            text123.Text = "A";

            run127.Append(runProperties127);
            run127.Append(text123);

            paragraph61.Append(paragraphProperties61);
            paragraph61.Append(run123);
            paragraph61.Append(run124);
            paragraph61.Append(run125);
            paragraph61.Append(run126);
            paragraph61.Append(run127);

            Paragraph paragraph62 = new Paragraph() { RsidParagraphMarkRevision = "0017479D", RsidParagraphAddition = "00FF3B05", RsidParagraphProperties = "00FF3B05", RsidRunAdditionDefault = "00FF3B05" };

            ParagraphProperties paragraphProperties62 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId22 = new ParagraphStyleId() { Val = "1" };

            ParagraphMarkRunProperties paragraphMarkRunProperties61 = new ParagraphMarkRunProperties();
            Bold bold65 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript71 = new BoldComplexScript() { Val = false };
            //Color color24 = new Color() { Val = "auto" };
            FontSizeComplexScript fontSizeComplexScript182 = new FontSizeComplexScript() { Val = "22" };

            paragraphMarkRunProperties61.Append(bold65);
            paragraphMarkRunProperties61.Append(boldComplexScript71);
            //paragraphMarkRunProperties61.Append(color24);
            paragraphMarkRunProperties61.Append(fontSizeComplexScript182);

            paragraphProperties62.Append(paragraphStyleId22);
            paragraphProperties62.Append(paragraphMarkRunProperties61);

            Run run128 = new Run() { RsidRunProperties = "0017479D" };

            RunProperties runProperties128 = new RunProperties();
            Bold bold66 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript72 = new BoldComplexScript() { Val = false };
            //Color color25 = new Color() { Val = "auto" };
            FontSizeComplexScript fontSizeComplexScript183 = new FontSizeComplexScript() { Val = "22" };

            runProperties128.Append(bold66);
            runProperties128.Append(boldComplexScript72);
            //runProperties128.Append(color25);
            runProperties128.Append(fontSizeComplexScript183);
            Text text124 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text124.Text = "Ток потребления ";

            run128.Append(runProperties128);
            run128.Append(text124);

            Run run129 = new Run() { RsidRunProperties = "0017479D", RsidRunAddition = "004E6FEF" };

            RunProperties runProperties129 = new RunProperties();
            Bold bold67 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript73 = new BoldComplexScript() { Val = false };
            //Color color26 = new Color() { Val = "auto" };
            FontSizeComplexScript fontSizeComplexScript184 = new FontSizeComplexScript() { Val = "22" };

            runProperties129.Append(bold67);
            runProperties129.Append(boldComplexScript73);
            //runProperties129.Append(color26);
            runProperties129.Append(fontSizeComplexScript184);
            Text text125 = new Text();
            text125.Text = "ЭПУ";

            run129.Append(runProperties129);
            run129.Append(text125);

            Run run130 = new Run() { RsidRunProperties = "0017479D" };

            RunProperties runProperties130 = new RunProperties();
            Bold bold68 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript74 = new BoldComplexScript() { Val = false };
            //Color color27 = new Color() { Val = "auto" };
            FontSizeComplexScript fontSizeComplexScript185 = new FontSizeComplexScript() { Val = "22" };

            runProperties130.Append(bold68);
            runProperties130.Append(boldComplexScript74);
            //runProperties130.Append(color27);
            runProperties130.Append(fontSizeComplexScript185);
            Text text126 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text126.Text = " 2:";

            run130.Append(runProperties130);
            run130.Append(text126);

            Run run131 = new Run() { RsidRunAddition = "0017479D" };

            RunProperties runProperties131 = new RunProperties();
            Bold bold69 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript75 = new BoldComplexScript() { Val = false };
            //Color color28 = new Color() { Val = "auto" };
            FontSizeComplexScript fontSizeComplexScript186 = new FontSizeComplexScript() { Val = "22" };

            runProperties131.Append(bold69);
            runProperties131.Append(boldComplexScript75);
            //runProperties131.Append(color28);
            runProperties131.Append(fontSizeComplexScript186);
            Text text127 = new Text();
            text127.Text = "-----";

            run131.Append(runProperties131);
            run131.Append(text127);

            Run run132 = new Run() { RsidRunProperties = "0017479D", RsidRunAddition = "0017479D" };

            RunProperties runProperties132 = new RunProperties();
            Bold bold70 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript76 = new BoldComplexScript() { Val = false };
            //Color color29 = new Color() { Val = "auto" };
            FontSizeComplexScript fontSizeComplexScript187 = new FontSizeComplexScript() { Val = "22" };

            runProperties132.Append(bold70);
            runProperties132.Append(boldComplexScript76);
            //runProperties132.Append(color29);
            runProperties132.Append(fontSizeComplexScript187);
            TabChar tabChar5 = new TabChar();

            run132.Append(runProperties132);
            run132.Append(tabChar5);

            Run run133 = new Run() { RsidRunProperties = "0017479D" };

            RunProperties runProperties133 = new RunProperties();
            Bold bold71 = new Bold() { Val = false };
            BoldComplexScript boldComplexScript77 = new BoldComplexScript() { Val = false };
            //Color color30 = new Color() { Val = "auto" };
            FontSizeComplexScript fontSizeComplexScript188 = new FontSizeComplexScript() { Val = "22" };
            Languages languages47 = new Languages() { Val = "en-US" };

            runProperties133.Append(bold71);
            runProperties133.Append(boldComplexScript77);
            //runProperties133.Append(color30);
            runProperties133.Append(fontSizeComplexScript188);
            runProperties133.Append(languages47);
            Text text128 = new Text();
            text128.Text = "A";

            run133.Append(runProperties133);
            run133.Append(text128);

            paragraph62.Append(paragraphProperties62);
            paragraph62.Append(run128);
            paragraph62.Append(run129);
            paragraph62.Append(run130);
            paragraph62.Append(run131);
            paragraph62.Append(run132);
            paragraph62.Append(run133);


            // Close the handle explicitly.
            wordprocessingDocument.Close();
            
            

        }

        private void btnCreateDir_Click(object sender, EventArgs e)
        {
            string path = @"F:\" + cmbBts.Text + "_" + DateTime.Now.ToString("yyyy-MM-dd");

            try
            {
                if (Directory.Exists(path))
                {
                    MessageBox.Show("Directory already exist");
                }

                // Try to create the directory.
                DirectoryInfo dir = Directory.CreateDirectory(path);

                //Delete the directory.
                //dir.Delete(path);
                //Console.WriteLine("The directory was deleted successfully.");
            }
            catch (Exception)
            {
                MessageBox.Show("error1");
            }
            finally { }

            string sourcePath = @"F:\";
            string destinationPath = path;
            string sourceFileName = "Template.docx";
            string destinationFileName = "Учетная карточка.docx";
            string sourceFile = Path.Combine(sourcePath, sourceFileName);
            string destinationFile = Path.Combine(destinationPath, destinationFileName);

            File.Copy(sourceFile, destinationFile, true);


        }

        private void cmbBts_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
