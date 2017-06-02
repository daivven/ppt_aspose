using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Slides;
using Aspose.Slides.Export;


namespace ppttest
{
    public partial class Form1 : Form
    {
        public string dataDir = @"D:\【临时】\ppt_aspose\ppttest\ppttest\ppt\";
        public Form1()
        {            
            InitializeComponent();
        }

        private void btn_create_Click(object sender, EventArgs e)
        {
            Presentation pres = new Presentation(dataDir + "test.ppt");
            ISlide sld = pres.Slides[0];
            // sld.Name = "123456";

            // Add autoshape of rectangle type
            IShape shp1 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 40, 150, 50);
            shp1.AlternativeText = "User Defined";
            shp1.Name = "20170876";
            IShape shp2 = sld.Shapes.AddAutoShape(ShapeType.Moon, 160, 40, 150, 50);
            String alttext = "User Defined";
            int iCount = sld.Shapes.Count;
            for (int i = 0; i < iCount; i++)
            {
                AutoShape ashp = (AutoShape)sld.Shapes[i];
                if (String.Compare(ashp.AlternativeText, alttext, StringComparison.Ordinal) == 0)
                {
                    //ashp.Hidden = true;
                }
            }

            // Save presentation to disk
            pres.Save(dataDir + "Hiding_Shapes.pptx", SaveFormat.Ppt);
        }

        //Method implementation to find a shape in a slide using its alternative text
        public static IShape FindShape(ISlide slide, string alttext)
        {
            //Iterating through all shapes inside the slide
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                //If the alternative text of the slide matches with the required one then
                //return the shape
                if (slide.Shapes[i].AlternativeText.CompareTo(alttext) == 0)
                    return slide.Shapes[i];

            }
            return null;
        }


        private void btn_findshape_Click(object sender, EventArgs e)
        {
            using (Presentation p = new Presentation(dataDir + "FindingShapeInSlide.pptx"))
            {

                ISlide slide = p.Slides[0];
                //alternative text of the shape to be found
                IShape shape = FindShape(slide, "Shape1");
                if (shape != null)
                {
                    //Console.WriteLine("Shape Name: " + shape.Name);
                    MessageBox.Show(shape.Name);
                }
            }
        }

        private void btn_merge_Click(object sender, EventArgs e)
        {
            //using (Presentation destPres = new Presentation())
            //{
            //    using (Presentation srcPres = new Presentation(dataDir + "问题检查模板.pptx"))
            //    {

            //        //Instantiate ISlide from the collection of slides in source presentation along with
            //        //master slide
            //        ISlide SourceSlide = srcPres.Slides[0];
            //        IMasterSlide SourceMaster = SourceSlide.LayoutSlide.MasterSlide;

            //        //Clone the desired master slide from the source presentation to the collection of masters in the
            //        //destination presentation
            //        IMasterSlideCollection masters = destPres.Masters;
            //        IMasterSlide DestMaster = SourceSlide.LayoutSlide.MasterSlide;

            //        //Clone the desired master slide from the source presentation to the collection of masters in the
            //        //destination presentation
            //        IMasterSlide iSlide = masters.AddClone(SourceMaster);

            //        //Clone the desired slide from the source presentation with the desired master to the end of the
            //        //collection of slides in the destination presentation
            //        ISlideCollection slds = destPres.Slides;
            //        slds.AddClone(SourceSlide, iSlide, true);
            //        //Clone the desired master slide from the source presentation to the collection of masters in the //destination presentation
            //        //Save the destination presentation to disk

            //        destPres.Slides[0].Remove();
            //        destPres.Save(dataDir + "Aspose_out.pptx", SaveFormat.Pptx);
            //    }
            //}
            List<string> files = new List<string>();
            files.Add("20160727001.ppt");
            files.Add("20160727002.ppt");
            files.Add("20160727003.ppt");
            string mergeName = string.Empty;
            MergePPT(files,ref mergeName);
            if (!string.IsNullOrEmpty(mergeName))
            {
                MessageBox.Show("合并后的文件为：" + mergeName);
            }
        }

        /**************************
         如模板文件如为pptx格式,生成文件为ppt，生成后的ppt中的问题描述字体颜色会发生变化，因此需要将模板文件更改为ppt格式，但ppt格式中的
         隐藏shape.name会在生成的ppt中丢失，采取TextFrame.Text添加标记来区分是否是隐藏元素，并将需要存储的值存于TextFrame.Text中
         ***************************/
        private void btn_singlegenarate_Click(object sender, EventArgs e)
        {
            using (Presentation destPres = new Presentation())
            {
                using (Presentation srcPres = new Presentation(dataDir + "问题检查模板.ppt"))
                {

                    //Instantiate ISlide from the collection of slides in source presentation along with
                    //master slide
                    ISlide SourceSlide = srcPres.Slides[0];
                    IMasterSlide SourceMaster = SourceSlide.LayoutSlide.MasterSlide;

                    //Clone the desired master slide from the source presentation to the collection of masters in the
                    //destination presentation
                    IMasterSlideCollection masters = destPres.Masters;
                    IMasterSlide DestMaster = SourceSlide.LayoutSlide.MasterSlide;

                    //Clone the desired master slide from the source presentation to the collection of masters in the
                    //destination presentation
                    IMasterSlide iSlide = masters.AddClone(SourceMaster);

                    //Clone the desired slide from the source presentation with the desired master to the end of the
                    //collection of slides in the destination presentation
                    ISlideCollection slds = destPres.Slides;
                    slds.AddClone(SourceSlide, iSlide, true);
                    //Clone the desired master slide from the source presentation to the collection of masters in the //destination presentation
                    //Save the destination presentation to disk

                    destPres.Slides[0].Remove();

                    foreach (ISlide slide in slds)
                    {
                        //Iterating through all shapes inside the slide
                        for (int i = 0; i < slide.Shapes.Count; i++)
                        {
                            //If the alternative text of the slide matches with the required one then
                            //return the shape
                            //if (slide.Shapes[i].AlternativeText.CompareTo(alttext) == 0)
                            //    return slide.Shapes[i];

                            //if (slide.Shapes[i].AlternativeText == "检查类别")
                            //    slide.Shapes[i].AlternativeText = "123456";

                            if (slide.Shapes[i].Placeholder != null)
                            {
                                //Change the text of each placeholder
                                ((IAutoShape)slide.Shapes[i]).TextFrame.Text = "This is Placeholder";
                            }


                        }
                    }


                    ISlide sld = destPres.Slides[0];

                    //// Add an AutoShape of Rectangle type
                    //IAutoShape ashp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 50);

                    //// Remove any fill style associated with the AutoShape
                    //ashp.FillFormat.FillType = FillType.NoFill;
                    ////ashp.UseBackgroundFill = true;
                    ////ashp.LineFormat.Width = 0;
                    //ashp.LineFormat.FillFormat.FillType = FillType.NoFill;//设置无边框                
                    ////ashp.LineFormat.FillFormat.SolidFillColor.Color = Color.Red;                 

                    //// Access the TextFrame associated with the AutoShape
                    //ITextFrame tf = ashp.TextFrame;
                    ////tf.TextFrameFormat.
                    //tf.Text = "Aspose TextBox22";

                    //// Access the Portion associated with the TextFrame
                    //IPortion port = tf.Paragraphs[0].Portions[0];

                    //// Set the Font for the Portion
                    //port.PortionFormat.LatinFont = new FontData("Times New Roman");

                    //// Set Bold property of the Font
                    //port.PortionFormat.FontBold = NullableBool.True;

                    //// Set Italic property of the Font
                    //port.PortionFormat.FontItalic = NullableBool.True;

                    //// Set Underline property of the Font
                    //port.PortionFormat.FontUnderline = TextUnderlineType.Single;

                    //// Set the Height of the Font
                    //port.PortionFormat.FontHeight = 25;

                    //// Set the color of the Font
                    //port.PortionFormat.FillFormat.FillType = FillType.Solid;
                    //port.PortionFormat.FillFormat.SolidFillColor.Color = Color.Blue;

                    // ExEnd:SetTextFontProperties
                    // Write the PPTX to disk 
                    //presentation.Save(dataDir + "SetTextFontProperties.pptx", SaveFormat.Pptx);

                    IAutoShape ashp2 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 200, 50, 100, 50);

                    // Remove any fill style associated with the AutoShape
                    ashp2.FillFormat.FillType = FillType.NoFill;
                    ashp2.Hidden = true;
                    ashp2.Name = "code";
                    ashp2.AlternativeText = "hiddenfield";

                    // Access the TextFrame associated with the AutoShape
                    ITextFrame tf2 = ashp2.TextFrame;
                    tf2.Text = "#code#这是code值";

                    destPres.Save(dataDir + "single.ppt", SaveFormat.Ppt);
                }
            }
        }

        private void btn_findhidden_Click(object sender, EventArgs e)
        {
            using (Presentation p = new Presentation(dataDir + "single.ppt"))
            {

                ISlide slide = p.Slides[0];
                string code = string.Empty;

                //Iterating through all shapes inside the slide
                for (int i = 0; i < slide.Shapes.Count; i++)
                {

                    if (slide.Shapes[i].AsISlideComponent.GetType() == typeof(Aspose.Slides.AutoShape))
                    {
                        if (((IAutoShape)slide.Shapes[i]).TextFrame.Text == "检查类别")
                        {
                            ((IAutoShape)slide.Shapes[i]).TextFrame.Text = "修改后的检查类别";
                        }
                        if (((IAutoShape)slide.Shapes[i]).TextFrame.Text == "问题描述")
                        {
                            ((IAutoShape)slide.Shapes[i]).TextFrame.Text = "修改后的问题描述";
                        }
                        /*** 2003版本slide.Shapes[i]).Name值会丢失 
                        if (((IAutoShape)slide.Shapes[i]).Name == "code")
                        {
                       
                            code= ((IAutoShape)slide.Shapes[i]).TextFrame.Text;
                        }
                         ****/
                        if (((IAutoShape)slide.Shapes[i]).TextFrame.Text.Contains("#code#"))
                        {
                            code = ((IAutoShape)slide.Shapes[i]).TextFrame.Text.Replace("#code#", "");
                        }
                    }
                }
                p.Save(dataDir + "findhidden.ppt", SaveFormat.Ppt);
                MessageBox.Show(code);
            }
        }

        private void btn_mutigenerate_Click(object sender, EventArgs e)
        {
            bool result = CreatePPT("20160727001", "20160727001检查类别", "20160727001问题描述");
            result = CreatePPT("20160727002", "20160727002检查类别", "20160727002问题描述");
            result = CreatePPT("20160727003", "20160727003检查类别", "20160727003问题描述");
        }
        /// <summary>
        /// 根据模板生成ppt
        /// </summary>
        /// <param name="?"></param>
        /// <returns></returns>
        private bool CreatePPT(string code, string jclb, string wtms)
        {
            bool result = true;
            try
            {
                using (Presentation destPres = new Presentation())
                {
                    using (Presentation srcPres = new Presentation(dataDir + "问题检查模板.ppt"))
                    {

                        //Instantiate ISlide from the collection of slides in source presentation along with
                        //master slide
                        ISlide SourceSlide = srcPres.Slides[0];
                        IMasterSlide SourceMaster = SourceSlide.LayoutSlide.MasterSlide;

                        //Clone the desired master slide from the source presentation to the collection of masters in the
                        //destination presentation
                        IMasterSlideCollection masters = destPres.Masters;
                        IMasterSlide DestMaster = SourceSlide.LayoutSlide.MasterSlide;

                        //Clone the desired master slide from the source presentation to the collection of masters in the
                        //destination presentation
                        IMasterSlide iSlide = masters.AddClone(SourceMaster);

                        //Clone the desired slide from the source presentation with the desired master to the end of the
                        //collection of slides in the destination presentation
                        ISlideCollection slds = destPres.Slides;
                        slds.AddClone(SourceSlide, iSlide, true);
                        //Clone the desired master slide from the source presentation to the collection of masters in the //destination presentation
                        //Save the destination presentation to disk

                        destPres.Slides[0].Remove();

                        ISlide slide = destPres.Slides[0];

                        IAutoShape ashp = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 200, 50, 100, 50);

                        // Remove any fill style associated with the AutoShape
                        ashp.FillFormat.FillType = FillType.NoFill;
                        ashp.Hidden = true;
                        ashp.Name = "code";
                        ashp.AlternativeText = "hiddenfield";

                        // Access the TextFrame associated with the AutoShape
                        ITextFrame tf = ashp.TextFrame;
                        tf.Text = "#code#" + code;


                        //Iterating through all shapes inside the slide
                        for (int i = 0; i < slide.Shapes.Count; i++)
                        {

                            if (slide.Shapes[i].AsISlideComponent.GetType() == typeof(Aspose.Slides.AutoShape))
                            {
                                if (((IAutoShape)slide.Shapes[i]).TextFrame.Text == "检查类别")
                                {
                                    ((IAutoShape)slide.Shapes[i]).TextFrame.Text = jclb;
                                }
                                if (((IAutoShape)slide.Shapes[i]).TextFrame.Text == "问题描述")
                                {
                                    ((IAutoShape)slide.Shapes[i]).TextFrame.Text = wtms;
                                }
                                /*** 2003版本slide.Shapes[i]).Name值会丢失 
                                if (((IAutoShape)slide.Shapes[i]).Name == "code")
                                {
                       
                                    code= ((IAutoShape)slide.Shapes[i]).TextFrame.Text;
                                }
                                 ****/
                                if (((IAutoShape)slide.Shapes[i]).TextFrame.Text.Contains("#code#"))
                                {
                                    code = ((IAutoShape)slide.Shapes[i]).TextFrame.Text.Replace("#code#", "");
                                }
                            }
                        }

                        destPres.Save(dataDir + code + ".ppt", SaveFormat.Ppt);
                    }
                }
            }
            catch (Exception ex)
            {
                result = false;
            }
            return result;
        }

        /// <summary>
        /// 合并ppt：依次合并files中的ppt文件
        /// </summary>
        /// <param name="files">要合并的文件名列表</param>
        /// <param name="desPPT">合并后的文件名</param>
        private void MergePPT(List<string> files,ref string desPPT)
        {
            using (Presentation destPres = new Presentation(dataDir + files[0]))
            {
                for (int i = 1; i < files.Count; i++)
                {
                    using (Presentation srcPres = new Presentation(dataDir + files[i]))
                    {

                        //Instantiate ISlide from the collection of slides in source presentation along with
                        //master slide
                        ISlide SourceSlide = srcPres.Slides[0];
                        IMasterSlide SourceMaster = SourceSlide.LayoutSlide.MasterSlide;

                        //Clone the desired master slide from the source presentation to the collection of masters in the
                        //destination presentation
                        IMasterSlideCollection masters = destPres.Masters;
                        IMasterSlide DestMaster = SourceSlide.LayoutSlide.MasterSlide;

                        //Clone the desired master slide from the source presentation to the collection of masters in the
                        //destination presentation
                        IMasterSlide iSlide = masters.AddClone(SourceMaster);

                        //Clone the desired slide from the source presentation with the desired master to the end of the
                        //collection of slides in the destination presentation
                        ISlideCollection slds = destPres.Slides;
                        slds.AddClone(SourceSlide, iSlide, true);
                        //Clone the desired master slide from the source presentation to the collection of masters in the //destination presentation
                        //Save the destination presentation to disk
                    }
                }
                //destPres.Slides[0].Remove();
                desPPT = DateTime.Now.ToString("yyyyMMddhhmmsss") + ".ppt";
                destPres.Save(dataDir + desPPT, SaveFormat.Ppt);
            }
        }

        private void btn_split_Click(object sender, EventArgs e)
        {
            string message = string.Empty;
            bool result = SpiltPPT("20160727095953.ppt", ref message);
            if (result)
            {
                MessageBox.Show("拆分成功");
            }
            else
            {
                MessageBox.Show(message);
            }
        }

        private bool SpiltPPT(string file,ref string message)
        {
            //Aspose.Slides.License license = new Aspose.Slides.License();
            //license.SetLicense("Aspose.Slides.lic");
            bool result = true;
            try
            {
                using (Presentation srcPres = new Presentation(dataDir + file))
                { 
                    foreach (ISlide slide in srcPres.Slides)
                    {
                        using (Presentation destPres = new Presentation())
                        {                            
                            for (int i = 0; i < slide.Shapes.Count; i++)
                            {
                                if (slide.Shapes[i].AsISlideComponent.GetType() == typeof(Aspose.Slides.AutoShape))
                                {
                                    if (((IAutoShape)slide.Shapes[i]).TextFrame.Text.Contains("#code#"))
                                    {
                                        string code = ((IAutoShape)slide.Shapes[i]).TextFrame.Text.Replace("#code#", "");
                                        ISlideCollection slds = destPres.Slides;
                                        slds.AddClone(slide);
                                        destPres.Slides[0].Remove();//去掉空白页
                                        destPres.Save(dataDir + code + ".ppt", SaveFormat.Ppt);
                                    }
                                }                               
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result = false;
                message = "拆分ppt失败：" + ex.Message;
            }
            return result;
        }
    }
}
    
