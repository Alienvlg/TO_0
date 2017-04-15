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
            String Dir = @"F:\Учетная карточка1.docx";

            // Open a WordprocessingDocument for editing using the filepath.
            WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(Dir, true);
            
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
            Text text11 = new Text();
            text11.Text = "Экстренная информация:";
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















            // Close the handle explicitly.
            wordprocessingDocument.Close();
        }
    }
}
