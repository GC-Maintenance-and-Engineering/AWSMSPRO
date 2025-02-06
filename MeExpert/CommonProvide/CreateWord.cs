using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using MeExpert.Models;
using System.Windows.Forms;
using word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Net;
using DocumentFormat.OpenXml.Packaging;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Drawing;
using System.Web.Mvc;
using DocumentFormat.OpenXml.Wordprocessing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using System.Data.SqlClient;
using System.Data;
using System.Globalization;
using Microsoft.Office.Interop.Word;

namespace MeExpert.CommonProvide
{
    public class CreateWord
    {
        public static void CreateDoc(ProposalDoc data)
        {
            object path = System.IO.Path.Combine(HttpContext.Current.Server.MapPath("~/Word/"), "proposal_defautl_31102024.docx");
            object pathSaveAs = System.IO.Path.Combine(HttpContext.Current.Server.MapPath("~/Word/"), "proposal_change.docx");
            HttpContext.Current.Session.Timeout = 60;
            try
            {
               
                //HttpContext.Current.Response.Write("path:"+ path);
                object missing = Missing.Value;
                word.Application wordApp = new word.Application();
                word.Document doc = null;
                //HttpContext.Current.Response.Write("after word app");
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                doc = wordApp.Documents.Open(ref path, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref isVisible, ref missing, ref missing, ref missing, ref missing);
                doc.Activate();
                //foreach (word.Range range in doc.StoryRanges)
                //{
                    //cover
                    FindAndReplaceOnShape(doc, "[c_company]", data.company);
                    FindAndReplaceOnShape(doc, "[c_position]", data.position);
                    FindAndReplaceOnShape(doc, "[c_start_date]", data.start_date);
                    FindAndReplaceOnShape(doc, "[c_end_date]", data.end_date);
                    FindAndReplaceOnShape(doc, "[c_pps_date]", data.proposal_date);
                    FindAndReplaceOnShape(doc, "[c_rev_no]", "Rev."+data.rev_no);

                    //content
                    FindAndReplace(wordApp, "[start_date]", data.start_date);
                    FindAndReplace(wordApp, "[end_date]", data.end_date);
                    FindAndReplace(wordApp, "[pps_date]", data.proposal_date);
                    FindAndReplace(wordApp, "[rev_no]", data.rev_no);
                    FindAndReplace(wordApp, "[scope_of_service]", data.scope_service);
                    FindAndReplace(wordApp, "[period]", data.period);
                    FindAndReplace(wordApp, "[work_date]", data.working_date);
                    FindAndReplace(wordApp, "[position]", data.position);
                    FindAndReplace(wordApp, "[candidate_name]", data.candidate_name);
                    FindAndReplaceHeader(doc, "[pps_date]", data.proposal_date);
                    FindAndReplaceHeader(doc, "[rev_no]", "Rev."+data.rev_no);

                    //cal table
                    FindAndReplace(wordApp, "[act_salary]", String.Format("{0:n}", data.act_salary));
                    FindAndReplace(wordApp, "[mr_1_2]", String.Format("{0:n}", data.mr_1_2));
                    FindAndReplace(wordApp, "[mr_1_4]", String.Format("{0:n}", data.mr_1_4));
                    FindAndReplace(wordApp, "[mr_1_6]", String.Format("{0:n}", data.mr_1_6));
                    FindAndReplace(wordApp, "[mr_1_7]", String.Format("{0:n}", data.mr_1_7));
                    FindAndReplace(wordApp, "[tt_1_1]", String.Format("{0:n}", data.tt_1_1));
                    FindAndReplace(wordApp, "[tt_1_2]", String.Format("{0:n}", data.tt_1_2));
                    FindAndReplace(wordApp, "[tt_1_4]", String.Format("{0:n}", data.tt_1_4));
                    FindAndReplace(wordApp, "[tt_1_6]", String.Format("{0:n}", data.tt_1_6));
                    FindAndReplace(wordApp, "[tt_1_7]", String.Format("{0:n}", data.tt_1_7));
                    FindAndReplace(wordApp, "[ot_1_1]", String.Format("{0:n}", data.ot_1_1));
                    FindAndReplace(wordApp, "[mr_subtotal]", String.Format("{0:n}", data.mr_subtotal));
                    FindAndReplace(wordApp, "[subtotal1]", String.Format("{0:n}", data.subtotal1));
                    FindAndReplace(wordApp, "[tt_2_1]", String.Format("{0:n}", data.tt_2_1));
                    FindAndReplace(wordApp, "[tt_2_2]", String.Format("{0:n}", data.tt_2_2));
                    FindAndReplace(wordApp, "[tt_2_3]", String.Format("{0:n}", data.tt_2_3));
                    FindAndReplace(wordApp, "[subtotal2]", String.Format("{0:n}", data.subtotal2));
                    FindAndReplace(wordApp, "[subtotal2_dh]", String.Format("{0:n}", data.subtotal2_dh));
                    FindAndReplace(wordApp, "[grandtotal]", String.Format("{0:n}", data.grandtotal));
                    FindAndReplace(wordApp, "[grandtotal_dh]", String.Format("{0:n}", data.grandtotal_dh));
                //}
                doc.SaveAs(ref pathSaveAs, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                doc.Close(ref missing, ref missing, ref missing);
               // HttpContext.Current.Response.Write("end");
            }
            catch (Exception ex)
            {
                HttpContext.Current.Response.Write("error in catch"+ex.Message);
                throw ex;
            }
        }
        public static void FindAndReplace(word.Application wordApp, object findText, object replaceText)
        {
            try
            {
                object matchCase = true;
                object matchWholeWord = true;
                object matchwildCards = false;
                object matchSoundsLike = false;
                object matchAllWordForms = false;
                object forward = true;
                object format = false;
                object matchKashida = false;
                object matchDiacritics = false;
                object matchAlefHamza = false;
                object matchControl = false;
                object read_only = false;
                object visible = true;
                object replace = word.WdReplace.wdReplaceAll;
                object wrap = word.WdFindWrap.wdFindContinue;

                if (replaceText.ToString().Length < 256) // Normal execution
                {
                    wordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchwildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceText, ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
                }
                else  // Long string
                {
                    object missing = System.Reflection.Missing.Value;
                    wordApp.Selection.Find.Execute(
                    ref findText, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);

                    wordApp.Selection.Text = (string)replaceText;
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
           
            
        }
        public static void FindAndReplaceHeader(word.Document doc, object findText, object replaceText)
        {
            try
            {
                object replaceAll = word.WdReplace.wdReplaceAll;
                object missing = Missing.Value;
                foreach (word.Section section in doc.Sections)
                {
                    word.Range headerRange = section.Headers[word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Find.Text = (string)findText;
                    headerRange.Find.Replacement.Text = (string)replaceText;
                    headerRange.Find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
           
        }
        public static void FindAndReplaceOnShape(word.Document doc,string findText,string replaceText)
        {
            try
            {
                var range = doc.Range();
                range.Find.Execute(FindText: findText, Replace: word.WdReplace.wdReplaceAll, ReplaceWith: replaceText);

                var shapes = doc.Shapes;

                foreach (word.Shape shape in shapes)
                {
                    if (shape.TextFrame.HasText != 0)
                    {
                        var initialText = shape.TextFrame.TextRange.Text;
                        var resultingText = initialText.Replace(findText, replaceText);
                        shape.TextFrame.TextRange.Text = resultingText;
                    }
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
           

        }
    }
    public class CreateWordByOpenXML
    {
        public static MemoryStream CreateDoc(ProposalDoc data)
        {
            MemoryStream msResult = new MemoryStream();
            string path = System.IO.Path.Combine(HttpContext.Current.Server.MapPath("~/Word/"), "proposal_04112024.docx");
            byte[] byteArr = File.ReadAllBytes(path);
            MemoryStream stream = new MemoryStream(File.ReadAllBytes(path));
            //using (MemoryStream stream = new MemoryStream(File.ReadAllBytes(path)))
            //{
                stream.Write(byteArr, 0, (int)byteArr.Length);
                try
                {
                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
                    {
                    //header
                    var headers = wordDoc.MainDocumentPart.HeaderParts;
                    foreach (var headerPart in headers)
                    {
                        // Iterate through all text elements in the header
                        foreach (var text in headerPart.RootElement.Descendants<Text>())
                        {
                            if (text.Text.Contains("revno"))
                            {
                                // Replace the matched text
                                text.Text = text.Text.Replace("revno", data.rev_no);
                            }
                        }
                    }
                    //SearchAndReplaceHeader(wordDoc, "revno", data.rev_no);

                    //cover
                    SearchAndReplace(wordDoc, "c_company", data.company);
                    SearchAndReplace(wordDoc, "c_position", data.position);
                    SearchAndReplace(wordDoc, "c_start_date", data.start_date);
                    SearchAndReplace(wordDoc, "c_end_date", data.end_date);
                    //SearchAndReplace(wordDoc, "c_pps_date", data.proposal_date);

                    ////header
                    ////SearchAndReplaceHeader(wordDoc, "[proposal_date]", data.proposal_date);
                    //SearchAndReplaceHeader(wordDoc, "[rev_no]", data.proposal_date + " / Rev." + data.rev_no);

                    //content
                    SearchAndReplace(wordDoc, "start_date", data.start_date);
                    SearchAndReplace(wordDoc, "end_date", data.end_date);
                    //SearchAndReplace(wordDoc, "[pps_date]", data.proposal_date);
                    //SearchAndReplace(wordDoc, "[rev_no]", data.rev_no);
                    SearchAndReplace(wordDoc, "scope_of_service", data.scope_service);
                    SearchAndReplace(wordDoc, "salaryperiod", data.period);
                    SearchAndReplace(wordDoc, "work_date", data.working_date);
                    //SearchAndReplace(wordDoc, "can_position", data.position);
                    SearchAndReplace(wordDoc, "candidate_name", data.candidate_name);

                    //cal table
                    SearchAndReplace(wordDoc, "act_salary", String.Format("{0:n}", data.act_salary));
                    SearchAndReplace(wordDoc, "mr_1_2", String.Format("{0:n}", data.mr_1_2));
                    SearchAndReplace(wordDoc, "mr_1_4", String.Format("{0:n}", data.mr_1_4));
                    SearchAndReplace(wordDoc, "mr_1_6", String.Format("{0:n}", data.mr_1_6));
                    ////SearchAndReplace(wordDoc, "[mr_1_7]", String.Format("{0:n}", data.mr_1_7));
                    SearchAndReplace(wordDoc, "tt_1_1", String.Format("{0:n}", data.tt_1_1));
                    SearchAndReplace(wordDoc, "tt_1_2", String.Format("{0:n}", data.tt_1_2));
                    SearchAndReplace(wordDoc, "tt_1_4", String.Format("{0:n}", data.tt_1_4));
                    SearchAndReplace(wordDoc, "tt_1_6", String.Format("{0:n}", data.tt_1_6));
                    ////SearchAndReplace(wordDoc, "[tt_1_7]", String.Format("{0:n}", data.tt_1_7));
                    SearchAndReplace(wordDoc, "ot_1_1", String.Format("{0:n}", data.ot_1_1));
                    SearchAndReplace(wordDoc, "mr_subtotal", String.Format("{0:n}", data.mr_subtotal));
                    SearchAndReplace(wordDoc, "subtotal1", String.Format("{0:n}", data.subtotal1));
                    SearchAndReplace(wordDoc, "tt_2_1", String.Format("{0:n}", data.tt_2_1));
                    //SearchAndReplace(wordDoc, "[tt_2_2]", String.Format("{0:n}", data.tt_2_2));
                    SearchAndReplace(wordDoc, "tt_2_3", String.Format("{0:n}", data.tt_2_3));
                    SearchAndReplace(wordDoc, "tt_2_10", String.Format("{0:n}", data.tt_2_10));
                    //SearchAndReplace(wordDoc, "[tt_2_3]", String.Format("{0:n}", data.tt_2_3));
                    SearchAndReplace(wordDoc, "subtotal2", String.Format("{0:n}", data.subtotal2));
                    ////SearchAndReplace(wordDoc, "[subtotal2_dh]", String.Format("{0:n}", data.subtotal2_dh));
                    SearchAndReplace(wordDoc, "grandtotal", String.Format("{0:n}", data.grandtotal));
                    ////SearchAndReplace(wordDoc, "[grandtotal_dh]", String.Format("{0:n}", data.grandtotal_dh));

                    wordDoc.Close();
                        //wordDoc.MainDocumentPart.Document.Save();
                        //wordDoc.MainDocumentPart.FeedData(stream);
                    }
                    stream.Seek(0, SeekOrigin.Begin);
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                return stream;
            //}
         
        }
        public static void SearchAndReplace(WordprocessingDocument wordDoc, string searchText,object replaceText)
        {
            string docText = null;
            StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream());
            docText = sr.ReadToEnd();

            //Regex regex = new Regex([end_date]);
            docText = docText.Replace(searchText, replaceText.ToString());
            StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create));
            sw.Write(docText);
            sw.Flush();
        }
        public static void SearchAndReplaceHeader(WordprocessingDocument wordDoc, string searchText, object replaceText)
        {
           
                foreach (HeaderPart header in wordDoc.MainDocumentPart.HeaderParts)
                {
                    string docText = null;
                    StreamReader sr = new StreamReader(header.GetStream());
                    docText = sr.ReadToEnd();

                    //Regex regex = new Regex([end_date]);
                    docText = docText.Replace(searchText, replaceText.ToString());
                    StreamWriter sw = new StreamWriter(header.GetStream(FileMode.Create));
                    sw.Write(docText);
                    sw.Flush();
                }
        }

        public static List<FilePPS> GetDetailProposal(int proposal_id)
        { 
            ProposalDoc data = new ProposalDoc();
            Connect_DB db = new Connect_DB();
            SqlDataReader dr;
            SqlCommand cmd;
            string path = string.Empty;
            MemoryStream stream = new MemoryStream();
            string docFile = string.Empty;
            List<FilePPS> res = new List<FilePPS>();
            try
            {
                db.OpenConn();
                FilePPS filePPS = new FilePPS();
                string sql = $"select rev_no,p.file_name,start_date,end_date,working_date,proposal_date, " +
                    $"scope_service,act_salary,period,ps.position_name,c.client_name,d.first_name,d.last_name,pps_comp,pps_license from proposal_list p " +
                $" inner join req_success s on p.succ_id = s.succ_id" +
                $" inner join requisition r on s.req_id = r.req_id" +
                $" inner join position ps on r.position_id = ps.position_id" +
                $" inner join client c on r.client_id = c.client_id" +
                $" inner join candidate d on d.cand_id = s.cand_id" +
                $" where proposal_id = @proposal_id";
                cmd = new SqlCommand(sql, db.GetConn());
                cmd.Parameters.Add("@proposal_id", SqlDbType.Int).Value = proposal_id;
                dr = cmd.ExecuteReader();
                if (dr.Read())
                {
                  
                    data.company = $"{dr["client_name"]}";
                    data.position = $"{dr["position_name"]}";
                    data.rev_no = (dr["rev_no"] == DBNull.Value) ? string.Empty : $"{dr["rev_no"]}";
                    data.file_name = $"{dr["file_name"]}";
                    data.start_date = (dr["start_date"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{dr["start_date"]}", CultureInfo.CurrentCulture).ToString("MMMM' 'dd','yyyy");
                    data.end_date = (dr["end_date"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{dr["end_date"]}", CultureInfo.CurrentCulture).ToString("MMMM' 'dd','yyyy");
                    data.working_date = (dr["working_date"] == DBNull.Value) ? 0 : (int)dr["working_date"];
                    data.proposal_date = (dr["proposal_date"] == DBNull.Value) ? string.Empty : Convert.ToDateTime($"{dr["proposal_date"]}", CultureInfo.CurrentCulture).ToString("MMMM' 'dd','yyyy");
                    //data.scope_service = (dr["scope_service"] == DBNull.Value) ? string.Empty : $"{dr["scope_service"]}";
                    data.scope_service = "test";
                    data.act_salary = (dr["act_salary"] == DBNull.Value) ? 0 : (decimal)dr["act_salary"];
                    data.period = (dr["period"] == DBNull.Value) ? 0 : (int)dr["period"];
                    data.candidate_name = $"{dr["first_name"]} {dr["last_name"]}";
                    filePPS.position = data.position;
                    filePPS.rev_no = data.rev_no;
                    //on cal table
                    double work_price = 173.33; //working_date=5
                    if (data.working_date == 6)
                    {
                        work_price = 208;
                    }
                    data.mr_1_2 = (Decimal.ToDouble(data.act_salary) * 0.05);
                    data.mr_1_4 = (Decimal.ToDouble(data.act_salary) * 0.00134);
                    data.mr_1_6 = (Decimal.ToDouble(data.act_salary) * 0.006);
                    //data.mr_1_7 = (Decimal.ToDouble(data.act_salary) * 0.05);
                    data.tt_1_1 = Decimal.ToDouble(data.period) * Decimal.ToDouble(data.act_salary);
                    data.tt_1_2 = Decimal.ToDouble(data.period) * data.mr_1_2;
                    data.tt_1_4 = Decimal.ToDouble(data.period) * data.mr_1_4;
                    data.tt_1_6 = Decimal.ToDouble(data.period) * data.mr_1_6;
                    //data.tt_1_7 = Decimal.ToDouble(data.period) * data.mr_1_7;
                    data.ot_1_1 = (Decimal.ToDouble(data.act_salary) / work_price) * 1.05;
                    data.mr_subtotal = Decimal.ToDouble(data.act_salary) + data.mr_1_2 + 3919 + data.mr_1_4 + 2501 + data.mr_1_6 + 8146;
                    data.subtotal1 = data.tt_1_1 + data.tt_1_2 + data.tt_1_4 + data.tt_1_6 + 47028 + 30012 + 97752;
                    data.tt_2_1 = (data.period * work_price * 0.1) * (data.ot_1_1 * 1.5);
                    //data.tt_2_2 = Decimal.ToDouble(data.act_salary) * 4;
                    //data.tt_2_3 = Decimal.ToDouble(data.act_salary) * 3;
                    data.tt_2_3 = (dr["pps_comp"] == DBNull.Value) ? 0 : Decimal.ToDouble((decimal)dr["pps_comp"]) ;
                    data.tt_2_10 = (dr["pps_license"] == DBNull.Value) ? 0 : Decimal.ToDouble((decimal)dr["pps_license"]) ;
                    //data.grandtotal_dh = data.subtotal1 + data.subtotal2_dh;
                    data.subtotal2 = data.tt_2_1 + data.tt_2_10;
                    //data.subtotal2_dh = data.tt_2_1 + data.tt_2_3;
                    data.grandtotal = data.subtotal1 + data.subtotal2;
                    //data.grandtotal_dh = data.subtotal1 + data.subtotal2_dh;
                   
                }
                dr.Close();
                filePPS.stream = CreateDoc(data);
                res.Add(filePPS);
                Management.InsertLog($"Proposal", $"Download Proposal proposal_id={proposal_id}");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                //throw new Exception(ex.Message);
                //Console.WriteLine("Custom Error Text " + ex.Message);
            }
            finally
            {
                db.CloseConn();
            }
            return res;
        }
    }
}