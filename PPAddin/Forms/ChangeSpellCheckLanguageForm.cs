using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop;
namespace PPAddin.Forms
{
    public partial class ChangeSpellCheckLanguageForm : Form
    {
        public ChangeSpellCheckLanguageForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ChangeSpellCheckLanguage();
        }

        void ChangeSpellCheckLanguage()
        {

            string[] LanguageNames = new string[217];



            LanguageNames[1] = "Afrikaans";
            LanguageNames[2] = "Albanian";
            LanguageNames[3] = "Amharic";
            LanguageNames[4] = "Arabic";
            LanguageNames[5] = "Arabic Algeria";
            LanguageNames[6] = "Arabic Bahrain";
            LanguageNames[7] = "Arabic Egypt";
            LanguageNames[8] = "Arabic Iraq";
            LanguageNames[9] = "Arabic Jordan";
            LanguageNames[10] = "Arabic Kuwait";
            LanguageNames[11] = "Arabic Lebanon";
            LanguageNames[12] = "Arabic Libya";
            LanguageNames[13] = "Arabic Morocco";
            LanguageNames[14] = "Arabic Oman";
            LanguageNames[15] = "Arabic Qatar";
            LanguageNames[16] = "Arabic Syria";
            LanguageNames[17] = "Arabic Tunisia";
            LanguageNames[18] = "Arabic UAE";
            LanguageNames[19] = "Arabic Yemen";
            LanguageNames[20] = "Armenian";
            LanguageNames[21] = "Assamese";
            LanguageNames[22] = "Azerbaijani Cyrillic";
            LanguageNames[23] = "Azerbaijani Latin";
            LanguageNames[24] = "Basque (Basque)";
            LanguageNames[25] = "Belgian Dutch";
            LanguageNames[26] = "Belgian French";
            LanguageNames[27] = "Bengali";
            LanguageNames[28] = "Bosnian";
            LanguageNames[29] = "Bosnian Bosnia Herzegovina Cyrillic";
            LanguageNames[30] = "Bosnian Bosnia Herzegovina Latin";
            LanguageNames[31] = "Portuguese (Brazil)";
            LanguageNames[32] = "Bulgarian";
            LanguageNames[33] = "Burmese";
            LanguageNames[34] = "Belarusian";
            LanguageNames[35] = "Catalan";
            LanguageNames[36] = "Cherokee";
            LanguageNames[37] = "Chinese Hong Kong SAR";
            LanguageNames[38] = "Chinese Macao SAR";
            LanguageNames[39] = "Chinese Singapore";
            LanguageNames[40] = "Croatian";
            LanguageNames[41] = "Czech";
            LanguageNames[42] = "Danish";
            LanguageNames[43] = "Divehi";
            LanguageNames[44] = "Dutch";
            LanguageNames[45] = "Edo";
            LanguageNames[46] = "English AUS";
            LanguageNames[47] = "English Belize";
            LanguageNames[48] = "English Canadian";
            LanguageNames[49] = "English Caribbean";
            LanguageNames[50] = "English Indonesia";
            LanguageNames[51] = "English Ireland";
            LanguageNames[52] = "English Jamaica";
            LanguageNames[53] = "English NewZealand";
            LanguageNames[54] = "English Philippines";
            LanguageNames[55] = "English South Africa";
            LanguageNames[56] = "English Trinidad Tobago";
            LanguageNames[57] = "English UK";
            LanguageNames[58] = "English US";
            LanguageNames[59] = "English Zimbabwe";
            LanguageNames[60] = "Estonian";
            LanguageNames[61] = "Faeroese";
            LanguageNames[62] = "Farsi";
            LanguageNames[63] = "Filipino";
            LanguageNames[64] = "Finnish";
            LanguageNames[65] = "French";
            LanguageNames[66] = "French Cameroon";
            LanguageNames[67] = "French Canadian";
            LanguageNames[68] = "French Coted Ivoire";
            LanguageNames[69] = "French Haiti";
            LanguageNames[70] = "French Luxembourg";
            LanguageNames[71] = "French Mali";
            LanguageNames[72] = "French Monaco";
            LanguageNames[73] = "French Morocco";
            LanguageNames[74] = "French Reunion";
            LanguageNames[75] = "French Senegal";
            LanguageNames[76] = "French West Indies";
            LanguageNames[77] = "French Congo DRC";
            LanguageNames[78] = "Frisian Netherlands";
            LanguageNames[79] = "Fulfulde";
            LanguageNames[80] = "Irish (Ireland)";
            LanguageNames[81] = "Scottish Gaelic";
            LanguageNames[82] = "Galician";
            LanguageNames[83] = "Georgian";
            LanguageNames[84] = "German";
            LanguageNames[85] = "German Austria";
            LanguageNames[86] = "German Liechtenstein";
            LanguageNames[87] = "German Luxembourg";
            LanguageNames[88] = "Greek";
            LanguageNames[89] = "Guarani";
            LanguageNames[90] = "Gujarati";
            LanguageNames[91] = "Hausa";
            LanguageNames[92] = "Hawaiian";
            LanguageNames[93] = "Hebrew";
            LanguageNames[94] = "Hindi";
            LanguageNames[95] = "Hungarian";
            LanguageNames[96] = "Ibibio";
            LanguageNames[97] = "Icelandic";
            LanguageNames[98] = "Igbo";
            LanguageNames[99] = "Indonesian";
            LanguageNames[100] = "Inuktitut";
            LanguageNames[101] = "Italian";
            LanguageNames[102] = "Japanese";
            LanguageNames[103] = "Kannada";
            LanguageNames[104] = "Kanuri";
            LanguageNames[105] = "Kashmiri";
            LanguageNames[106] = "Kashmiri Devanagari";
            LanguageNames[107] = "Kazakh";
            LanguageNames[108] = "Khmer";
            LanguageNames[109] = "Kirghiz";
            LanguageNames[110] = "Konkani";
            LanguageNames[111] = "Korean";
            LanguageNames[112] = "Kyrgyz";
            LanguageNames[113] = "Lao";
            LanguageNames[114] = "Latin";
            LanguageNames[115] = "Latvian";
            LanguageNames[116] = "Lithuanian";
            LanguageNames[117] = "Macedonian FYROM";
            LanguageNames[118] = "Malayalam";
            LanguageNames[119] = "Malay Brunei Darussalam";
            LanguageNames[120] = "Malaysian";
            LanguageNames[121] = "Maltese";
            LanguageNames[122] = "Manipuri";
            LanguageNames[123] = "Maori";
            LanguageNames[124] = "Marathi";
            LanguageNames[125] = "Mexican Spanish";
            LanguageNames[126] = "Mixed";
            LanguageNames[127] = "Mongolian";
            LanguageNames[128] = "Nepali";
            LanguageNames[129] = "No specified";
            LanguageNames[130] = "No proofing";
            LanguageNames[131] = "Norwegian Bokmol";
            LanguageNames[132] = "Norwegian Nynorsk";
            LanguageNames[133] = "Odia";
            LanguageNames[134] = "Oromo";
            LanguageNames[135] = "Pashto";
            LanguageNames[136] = "Polish";
            LanguageNames[137] = "Portuguese";
            LanguageNames[138] = "Punjabi";
            LanguageNames[139] = "Quechua Bolivia";
            LanguageNames[140] = "Quechua Ecuador";
            LanguageNames[141] = "Quechua Peru";
            LanguageNames[142] = "Rhaeto Romanic";
            LanguageNames[143] = "Romanian";
            LanguageNames[144] = "Romanian Moldova";
            LanguageNames[145] = "Russian";
            LanguageNames[146] = "Russian Moldova";
            LanguageNames[147] = "Sami Lappish";
            LanguageNames[148] = "Sanskrit";
            LanguageNames[149] = "Sepedi";
            LanguageNames[150] = "Serbian Bosnia Herzegovina Cyrillic";
            LanguageNames[151] = "Serbian Bosnia Herzegovina Latin";
            LanguageNames[152] = "Serbian Cyrillic";
            LanguageNames[153] = "Serbian Latin";
            LanguageNames[154] = "Sesotho";
            LanguageNames[155] = "Simplified Chinese";
            LanguageNames[156] = "Sindhi";
            LanguageNames[157] = "Sindhi Pakistan";
            LanguageNames[158] = "Sinhalese";
            LanguageNames[159] = "Slovak";
            LanguageNames[160] = "Slovenian";
            LanguageNames[161] = "Somali";
            LanguageNames[162] = "Sorbian";
            LanguageNames[163] = "Spanish";
            LanguageNames[164] = "Spanish Argentina";
            LanguageNames[165] = "Spanish Bolivia";
            LanguageNames[166] = "Spanish Chile";
            LanguageNames[167] = "Spanish Colombia";
            LanguageNames[168] = "Spanish Costa Rica";
            LanguageNames[169] = "Spanish Dominican Republic";
            LanguageNames[170] = "Spanish Ecuador";
            LanguageNames[171] = "Spanish El Salvador";
            LanguageNames[172] = "Spanish Guatemala";
            LanguageNames[173] = "Spanish Honduras";
            LanguageNames[174] = "Spanish Modern Sort";
            LanguageNames[175] = "Spanish Nicaragua";
            LanguageNames[176] = "Spanish Panama";
            LanguageNames[177] = "Spanish Paraguay";
            LanguageNames[178] = "Spanish Peru";
            LanguageNames[179] = "Spanish Puerto Rico";
            LanguageNames[180] = "Spanish Uruguay";
            LanguageNames[181] = "Spanish Venezuela";
            LanguageNames[182] = "Sutu";
            LanguageNames[183] = "Swahili";
            LanguageNames[184] = "Swedish";
            LanguageNames[185] = "Swedish Finland";
            LanguageNames[186] = "Swiss French";
            LanguageNames[187] = "Swiss German";
            LanguageNames[188] = "Swiss Italian";
            LanguageNames[189] = "Syriac";
            LanguageNames[190] = "Tajik";
            LanguageNames[191] = "Tamazight";
            LanguageNames[192] = "Tamazight Latin";
            LanguageNames[193] = "Tamil";
            LanguageNames[194] = "Tatar";
            LanguageNames[195] = "Telugu";
            LanguageNames[196] = "Thai";
            LanguageNames[197] = "Tibetan";
            LanguageNames[198] = "Tigrigna Eritrea";
            LanguageNames[199] = "Tigrigna Ethiopic";
            LanguageNames[200] = "Traditional Chinese";
            LanguageNames[201] = "Tsonga";
            LanguageNames[202] = "Tswana";
            LanguageNames[203] = "Turkish";
            LanguageNames[204] = "Turkmen";
            LanguageNames[205] = "Ukrainian";
            LanguageNames[206] = "Urdu";
            LanguageNames[207] = "Uzbek Cyrillic";
            LanguageNames[208] = "Uzbek Latin";
            LanguageNames[209] = "Venda";
            LanguageNames[210] = "Vietnamese";
            LanguageNames[211] = "Welsh";
            LanguageNames[212] = "Xhosa";
            LanguageNames[213] = "Yi";
            LanguageNames[214] = "Yiddish";
            LanguageNames[215] = "Yoruba";
            LanguageNames[216] = "Zulu";


            int[] LanguageIDs = new int[] {
            -1,1078,1052,1118,1025,5121,15361,3073,2049,11265,13313,12289,4097,6145,8193,16385,10241,7169,14337,9217,1067,1101,2092,1068,1069,2067,2060,1093,4122,8218,5146,1046,1026,1109,1059,1027,1116,3076,5124,4100,1050,1029,1030,1125,1043,1126,3081,10249,4105,9225,14345,6153,8201,5129,13321,7177,11273,2057,1033,12297,1061,1080,1065,1124,1035,1036,11276,3084,12300,15372,5132,13324,6156,14348,8204,10252,7180,9228,1122,1127,2108,1084,1110,1079,1031,3079,5127,4103,1032,1140,1095,1128,1141,1037,1081,1038,1129,1039,1136,1057,1117,1040,1041,1099,1137,1120,2144,1087,1107,1088,1111,1042,1088,1108,1142,1062,1063,1071,1100,2110,1086,1082,1112,1153,1102,2058,-2,1104,1121,0,1024,1044,2068,1096,1138,1123,1045,2070,1094,1131,2155,3179,1047,1048,2072,1049,2073,1083,1103,1132,7194,6170,3098,2074,1072,2052,1113,2137,1115,1051,1060,1143,1070,1034,11274,16394,13322,9226,5130,7178,12298,17418,4106,18442,3082,19466,6154,15370,10250,20490,14346,8202,1072,1089,1053,2077,4108,2055,2064,1114,1064,1119,2143,1097,1092,1098,1054,1105,2163,1139,1028,1073,1074,1055,1090,1058,1056,2115,1091,1075,1066,1106,1076,1144,1085,1130,1077
            };

            MessageBox.Show(LanguageIDs.Length.ToString());
            this.Hide();

            string TargetLanguageID;
            TargetLanguageID = LanguageIDs[ComboBox1.SelectedIndex + 1].ToString();


            string TargetLanguage;
            TargetLanguage = LanguageNames[ComboBox1.SelectedIndex + 1];


            Microsoft.Office.Interop.PowerPoint.Slide PresentationSlide;
            Microsoft.Office.Interop.PowerPoint.Shape SlideShape;
            SmartArtNode SlideSmartArtNode;
            int GroupCount;


            //# If Mac Then
            //        'Mac does not (yet) support property .HasHandoutMaster


            //On Error Resume Next
            //For Each SlideShape In ActivePresentation.HandoutMaster.Shapes
            //    If SlideShape.HasTextFrame Then
            //    SlideShape.TextFrame2.TextRange.LanguageID = TargetLanguageID
            //    End If
            //Next
            //On Error GoTo 0

            //# Else

            if (ThisAddIn.application.ActivePresentation.HasHandoutMaster) {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape sh in ThisAddIn.application.ActivePresentation.HandoutMaster.Shapes) {
                    if (sh.HasTextFrame == MsoTriState.msoCTrue || sh.HasTextFrame == MsoTriState.msoTrue)
                        sh.TextFrame2.TextRange.LanguageID = (MsoLanguageID)(int.Parse(TargetLanguageID));

                }
            }

            //# End If
            if (ThisAddIn.application.ActivePresentation.HasTitleMaster == MsoTriState.msoTrue || ThisAddIn.application.ActivePresentation.HasTitleMaster == MsoTriState.msoCTrue)
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape sh in ThisAddIn.application.ActivePresentation.TitleMaster.Shapes) {
                    if (sh.HasTextFrame == MsoTriState.msoCTrue || sh.HasTextFrame == MsoTriState.msoTrue)
                    {
                        sh.TextFrame2.TextRange.LanguageID = (MsoLanguageID)int.Parse(TargetLanguageID);
                    }
                }
            }

            //# If Mac Then
            //        'Mac does not (yet) support property .HasNotesMaster


            //On Error Resume Next
            //For Each SlideShape In ActivePresentation.NotesMaster.Shapes
            //    If SlideShape.HasTextFrame Then
            //    SlideShape.TextFrame2.TextRange.LanguageID = TargetLanguageID
            //    End If
            //Next
            //On Error GoTo 0

            //# Else
            if (ThisAddIn.application.ActivePresentation.HasNotesMaster)
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape sh in ThisAddIn.application.ActivePresentation.NotesMaster.Shapes)
                {
                    if (sh.HasTextFrame == MsoTriState.msoCTrue || sh.HasTextFrame == MsoTriState.msoTrue)
                    {
                        sh.TextFrame2.TextRange.LanguageID = (MsoLanguageID)int.Parse(TargetLanguageID);
                    }
                }
            }


            //# End If
            ProgressForm pf = new ProgressForm();
            pf.Show();

            foreach (Microsoft.Office.Interop.PowerPoint.Slide ps in ThisAddIn.application.ActivePresentation.Slides)
            {
                pf.SetProgress((ps.SlideNumber * 100) / ThisAddIn.application.ActivePresentation.Slides.Count);
                foreach (Microsoft.Office.Interop.PowerPoint.Shape ss in ps.Shapes)
                {
                    ChangeShapeSpellCheckLanguage(ss, TargetLanguageID);
                }
            }





            pf.Hide();
            foreach (Microsoft.Office.Interop.PowerPoint.Shape ss in ThisAddIn.application.ActivePresentation.SlideMaster.Shapes)
            {
                ChangeShapeSpellCheckLanguage(ss, TargetLanguageID);
            }


            MessageBox.Show("Changed spellcheck language to " + TargetLanguage + " on all slides.");



        }
        void ChangeShapeSpellCheckLanguage(Microsoft.Office.Interop.PowerPoint.Shape  SlideShape,string  TargetLanguageID)
        {
            if(SlideShape.Type ==MsoShapeType.msoGroup)
            {
                var SlideShapeGroup = SlideShape.GroupItems;
                foreach (Microsoft.Office.Interop.PowerPoint.Shape SlideShapeChild in SlideShapeGroup)
                {
                    ChangeShapeSpellCheckLanguage (SlideShapeChild, TargetLanguageID);
                }
            }
            else
            {
                if (SlideShape.HasTextFrame==MsoTriState.msoCTrue|| SlideShape.HasTextFrame == MsoTriState.msoTrue) {
                    SlideShape.TextFrame2.TextRange.LanguageID = (MsoLanguageID)int.Parse(TargetLanguageID);
                 }


                if (SlideShape.HasTable==MsoTriState.msoCTrue || SlideShape.HasTable == MsoTriState.msoTrue) {
                    for (int TableRow = 1; TableRow <= SlideShape.Table.Rows.Count; TableRow++) {
                        for (int TableColumn = 1; TableColumn <= SlideShape.Table.Columns.Count; TableColumn++) {
                            SlideShape.Table.Cell(TableRow, TableColumn).Shape.TextFrame2.TextRange.LanguageID =(MsoLanguageID ) int.Parse(TargetLanguageID);
                    }
                      }
                  }

                if (SlideShape.HasSmartArt==MsoTriState.msoCTrue|| SlideShape.HasSmartArt == MsoTriState.msoTrue) {


                    for (int SlideShapeSmartArtNode = 1; SlideShapeSmartArtNode <= SlideShape.SmartArt.AllNodes.Count; SlideShapeSmartArtNode++) {

                        foreach (SmartArtNode SlideSmartArtNode in SlideShape.SmartArt.AllNodes) {
                            SlideSmartArtNode.TextFrame2.TextRange.LanguageID = (MsoLanguageID)int.Parse(TargetLanguageID);
                 }

               }
        }


            }
        }
      
    }
}
