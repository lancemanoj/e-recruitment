				// 25 SEP 23 LJL added TLS protocol to fix PDF output problem for VNs that occurred for days.
				//System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
				// USED " + ex.Message + " to find the exact error
				//error: The underlying connection was closed: An unexpected error occurred on a send
				//https://learn.microsoft.com/en-us/answers/questions/362284/error-the-underlying-connection-was-closed-an-unex
				
				//08 OCT 24 LJL revised some of the output text for front page and then looked over code.  need to know where front page and individual logo image is coming from as it doesn't show in DEMO, which is ok
				//08 OCT 24 LJL added separate VN Title and VN number to show on sep lines in the PDF front page
				// 09 OCT 24 LJL added diversity to see if it works
				// ' 20 OCT 24 LJL modified addDots for each TOC item - two pages 
//30 dec 24 manoj added code to check issue logo not appear in demo
				
using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.IO;
using System.Diagnostics;
using iTextSharp.text;
using iTextSharp.text.html;
using iTextSharp.text.pdf;
using iTextSharp.text.html.simpleparser;
using System.util;
using System.Net;
using System.Xml;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using System.Text.RegularExpressions;

public partial class GeneratePDF : System.Web.UI.Page
{
    Document _doc;
    
    MemoryStream _output = new MemoryStream();
    PdfWriter _pdfWriter;
    private List<TableOfContentsEntry> _contentsTable;
	private List<TableOfContentsEntry> _contentsTable1;
    Int32 _pageCount = int.MinValue;

	static bool _isAppNameNeed;
    static string _globalValue;
    static string _VN;
// 08 OCT 24 LJL add VN number and Title
	static string _VNNumber;
	static string _VNTitle;

    static string _UseToc;
    static bool _ShowTocPageNumbers = true;  
	
    static int _MaxIndexPage=1;
    static int _MinIndexPage=int.MinValue;
	static int _totalpages;
	static int _tocpagesCount;
	static string _vNClosing;
	static string _org;
	static string _orgCode;
	static string _ApplicantName;
	float TEXTSIZE = 10;
    /// <summary>
    /// Access routine for global variable.
    /// </summary>
	
	public static int TocPagesCount
    {
		get { return _tocpagesCount; }
        set { _tocpagesCount = value; }
    }	
	public static int TotalPages
    {
		get { return _totalpages; }
        set { _totalpages = value; }
    }
	public static string VNClosing
    {
		get { return _vNClosing; }
        set { _vNClosing = value; }
    }	
	public static string Org
    {
		get { return _org; }
        set { _org = value; }
    }
	public static string OrgCode
    {
		get { return _orgCode; }
        set { _orgCode = value; }
    }
	public static string ApplicantName
    {
		get { return _ApplicantName; }
        set { _ApplicantName = value; }
    }
    public static string GlobalValue
    {
        get { return _globalValue; }
        set { _globalValue = value; }
    } 
	public static bool IsAppNameNeed
    {
        get { return _isAppNameNeed; }
        set { _isAppNameNeed = value; }
    }
    public static string VN
    {
        get { return _VN; }
        set { _VN = value; }
    }
// 08 OCT 24 LJL add VN number and Title
	public static string VNNumber
	{
		get { return _VNNumber; }
		set { _VNNumber = value; }
	}

	public static string VNTitle
	{
		get { return _VNTitle; }
		set { _VNTitle = value; }
	}


    public static string UseToc
    {
        get { return _UseToc; }
        set { _UseToc = value; }
    }  

    public static bool ShowTocPageNumbers
    {
        get { return _ShowTocPageNumbers; }
        set { _ShowTocPageNumbers = value; }
    } 
	
    public static int MinIndexPage 
    {
        get { return _MinIndexPage; }
        set { _MinIndexPage = value; }
    }
    public static int MaxIndexpage
    {
        get { return _MaxIndexPage; }
        set { _MaxIndexPage = value; }
    }	
	
	private int timeOut;
	
	override protected void OnInit(EventArgs e)
	{		
		base.OnInit(e);
		
		timeOut = Server.ScriptTimeout;
		Server.ScriptTimeout = 3600;
		
	}
	protected override void OnUnload(EventArgs e)
    {
        base.OnUnload(e);

        Server.ScriptTimeout = timeOut;
    }

    protected void Page_Load(object sender, EventArgs e)
    {	
			
		int repeateCount = 2;
		TotalPages = 0;
		GlobalValue = "";
		TEXTSIZE = 10;
		bool InfoPage = false;
		bool tocNumbers = true;
		int viewFormat = 1;
		String dest_dir = "E:\\docs\\UNSHARE\\PDFmake\\PDF-CV\\";	
		String imageFolder = "E:\\docs\\UNSHARE\\vac-cv\\";	
		String deleteOldVNFiles = "E:\\docs\\UNSHARE\\PDFmake\\PDF-VN\\";
		if(Request.Url.Host.ToLower().Contains("wto"))
		{
			dest_dir = "E:\\docs\\WTO\\PDFmake\\PDF-CV\\";
			imageFolder = "E:\\docs\\WTO\\vac-cv\\";	
			deleteOldVNFiles = "E:\\docs\\WTO\\PDFmake\\PDF-VN\\";
		}
		if(Request.Url.ToString().ToLower().Contains("/demo/"))
		{
			dest_dir = "E:\\docs\\stage\\ALLORG\\PDFmake\\PDF-CV\\";	
			deleteOldVNFiles = "E:\\docs\\stage\\ALLORG\\PDFmake\\PDF-VN\\";				
		    imageFolder = "E:\\docs\\stage\\ALLORG\\vac-cv\\";	
		}	
		//Response.Write(Request.Url.ToString().ToLower()); return;				
		
		#region Delete old PDF files 
		string[] filePaths = Directory.GetFiles(dest_dir);
		foreach (string filePath in filePaths)
		{
			string extension = Path.GetExtension(filePath);
			if(File.GetCreationTime(filePath) < DateTime.Now.AddMinutes(-30) && (extension == ".pdf" ))//|| extension == ".html"
			{
				try
				{ 
					File.Delete(filePath);
				}
				catch (Exception ex)
				{
					//Response.Write(ex.ToString());
				}
			}
		}
		
		filePaths = Directory.GetFiles(deleteOldVNFiles);
		foreach (string filePath in filePaths)
		{
			string extension = Path.GetExtension(filePath);
			if(File.GetCreationTime(filePath) < DateTime.Now.AddMinutes(-30) && (extension == ".pdf" || extension == ".html"))//
			{
				try
				{ 
					File.Delete(filePath);
				}
				catch (Exception ex)
				{
					//Response.Write(ex.ToString());
				}
			}
		}				
		#endregion
		
		
  
		//For Admin Pdf output
		if (Request.QueryString["OutputType"] != null && Request.QueryString["OutputType"] == "Admin")
		{			
			try
			{             
				#region Get Html File From Query String
				
				String htmlfile = String.Empty;
				String clr = String.Empty;
				String lng = String.Empty;
				ApplicantName = String.Empty;
				string A4Size= string.Empty;				
				int usepix = 0;
				MinIndexPage = int.MinValue;
				MaxIndexpage = 1;
				bool multi = false;
				int candsCount = 0;
				if (Request.QueryString["htmlfile"] != null)
				{
					htmlfile = (Request.QueryString["htmlfile"])+ ".html";
					if(htmlfile.Contains("..\\"))
					{
						htmlfile = htmlfile.Replace("..\\", "");
					}
					else if(htmlfile.Contains("\\"))
					{
						htmlfile = htmlfile.Replace("\\", "");
					}
					
					//clr = clr = "#4000FF"; (Request.QueryString["clr"]);
					clr = "#"+Request.QueryString["clr"];
					VN = Request.QueryString["Vnname"];
					
					// 08 OCT 24 LJL added number and title of VN
					VNTitle = Request.QueryString["VNTitle"];
					VNNumber = Request.QueryString["VNNumber"];

					UseToc = Request.QueryString["UseToc"];
					A4Size = Request.QueryString["ASize"];
					VNClosing = Request.Params["VNClosing"];
					Org = Request.QueryString["Org"];
					OrgCode = Request.QueryString["OrgCode"];
					ApplicantName = Regex.Replace(Request.QueryString["ApplicantName"], "[^a-zA-Z]", "");
					lng = Request.QueryString["Lng"];
					viewFormat = String.IsNullOrEmpty(Request.QueryString["ViewFormat"]) ? 1 : Convert.ToInt32(Request.QueryString["ViewFormat"]);
					usepix = String.IsNullOrEmpty(Request.QueryString["Usepix"]) ? 0 : Convert.ToInt32(Request.QueryString["Usepix"]);
					TEXTSIZE = String.IsNullOrEmpty(Request.QueryString["TextSize"]) ? TEXTSIZE : Convert.ToSingle(Request.QueryString["TextSize"]);
					InfoPage = String.IsNullOrEmpty(Request.QueryString["InfoPage"]) ? InfoPage : (Request.QueryString["InfoPage"] == "Yes" ? true : false);
					if (htmlfile.Contains("MULTI")) multi = true;
					
					if(multi)
						int.TryParse(Request.QueryString["candsCount"], out candsCount);
					tocNumbers = String.IsNullOrEmpty(Request.QueryString["tocnumbers"]) ? tocNumbers : (Request.QueryString["tocnumbers"] == "Yes" ? true : false);
				}
				
				IsAppNameNeed = false;	
				
				string urlImgLogo = (Request.Url.ToString().ToLower().Contains("/demo/") ? "http" : Request.Url.Scheme) + Uri.SchemeDelimiter +Request.Url.Host + "/css/" + OrgCode + "-css/Logo-Pdf-Print.jpg";
				
				ShowTocPageNumbers = tocNumbers;
				
				string stringWithHTMLTags = "";
				
				try	
				{
					stringWithHTMLTags = File.ReadAllText(dest_dir + htmlfile,Encoding.Default);
				}
				catch(DirectoryNotFoundException ex)
				{
					ErrorHendler(ex, Request.Url.ToString());
					
					Response.Write("File does not exist - level 5<br>");
					Response.Flush();
					Response.End();
				}
				catch(FileNotFoundException ex)
				{
					//ErrorHendler(ex, Request.Url.ToString());
					
					Response.Write("File does not exist - level 6<br>");
					Response.Flush();
					Response.End();
				}

				if(usepix == 1)
					stringWithHTMLTags = stringWithHTMLTags.Replace("<img src=\"Doc", "<img src=\"" + imageFolder + "Doc");
				
				List<IElement> elements = new List<IElement>();
				try	
				{
					elements = ProcessHTML(stringWithHTMLTags);
				}
				catch(Exception ex)				
				{
					ErrorHendler(ex, Request.Url.ToString());
					
					string errormsg = "<br><br><br><div align='center' ><h1 style='color:red;'>Error occurred while generating the PDF</h1><br><table width='90%' >"+
									"<tr><td valign='top' bgcolor='#FFFBDF' colspan=2><font>This item has been logged with the Support Team and will be resolved soon.<br><Br>Should the error result in delay in your application, it will be noted and considered."+
									"<br>Otherwise, please allow 12 hours for non-essential issues to be resolved.<br></font></td></tr></table></div>";
					
					elements = ProcessHTML(errormsg);			
					
				}
				

				string pdffilename = htmlfile.Split('.')[0] + ".pdf";
				if(!multi)
					pdffilename = pdffilename.Replace("ADMIN_CV_", "ADMIN_CV_" + ApplicantName + "_");
				//else
					//pdffilename += lng;	
					
				#endregion		
	
				
				//we have to repeate because we want to keep internal links and bookmarks, also to have right numbering for pages and TOC
				for(int cp = 0; cp < repeateCount; cp++)
				{
					if (A4Size == "A4")
					{
						_doc = new Document(PageSize.A4, 30, 15, 25, 25);
					}
					else if (A4Size == "Letter")
					{
						_doc = new Document(PageSize.LETTER, 30, 15, 25, 25);
					}
					else
					{
						_doc = new Document();
					}	

					_pdfWriter = PdfWriter.GetInstance(_doc, new FileStream(dest_dir + pdffilename, FileMode.Create));
					var PageMode = PdfWriter.PageLayoutTwoColumnLeft;
					if(viewFormat == 0)
						PageMode = PdfWriter.PageLayoutOneColumn;
						
					_pdfWriter.ViewerPreferences = PageMode;
					
					//Step2 After that make certain you set linear page mode, which ensures that page tree has a linear structure and allows you to insert pages in the middle of the document and  re-order them.
					_pdfWriter.SetLinearPageMode();
					
					 OpenDocument();
					 String Test = String.Empty;
					 int Cont = 0;
					 _pdfWriter.PageEvent = new PageEventHelper();
					 
					if(multi)
					{					
						_doc.NewPage();
						Paragraph First = new Paragraph();
						First.Add(new Chunk("\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n"));
						_doc.Add(First);
						
						try{
							iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(new Uri(urlImgLogo));
							jpg.ScaleToFit(50f, 50f);
							jpg.Alignment = iTextSharp.text.Image.TEXTWRAP ;
							jpg.SetAbsolutePosition(_doc.PageSize.Width/2 - 15f,_doc.PageSize.Height/1.5f - 10f);
							_doc.Add(jpg);
						}
						catch{}	
					}
					else if(InfoPage)
					{
						_doc.NewPage();
						Paragraph First = new Paragraph();
						First.Add(new Chunk("\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n"));
						_doc.Add(First);
						
						try{
							iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(new Uri(urlImgLogo));
							jpg.ScaleToFit(50f, 50f);
							jpg.Alignment = iTextSharp.text.Image.TEXTWRAP ;
							jpg.SetAbsolutePosition(_doc.PageSize.Width/2 - 15f,_doc.PageSize.Height/1.5f - 10f);
							_doc.Add(jpg);
						}
						catch{}	
					}
					
					if(VNClosing != null && VNClosing.Trim() != "")
					{						
						DateTime parsedDate = DateTime.Now;
						try
						{
							parsedDate = DateTime.ParseExact(VNClosing, "d", new CultureInfo("en-US"));						
						}
						catch
						{
							//parsedDate = DateTime.ParseExact(VNClosing, "dd MM yy", new CultureInfo("en-US"));
							try
							{
								parsedDate = DateTime.ParseExact(VNClosing, "dd MM yy", new CultureInfo("en-US"));
							}
							catch
							{
								parsedDate = Convert.ToDateTime(VNClosing);
							}
						}
						Paragraph vnParagraph = new Paragraph(VNTitle + Environment.NewLine + Environment.NewLine, FontFactory.GetFont("Arial", 18, Font.BOLD));
						vnParagraph.Alignment = Element.ALIGN_CENTER;
						_doc.Add(vnParagraph);
						
	Paragraph vnNumberParagraph = new Paragraph(VNNumber + Environment.NewLine + Environment.NewLine, FontFactory.GetFont("Arial", 17, Font.BOLD));
	vnNumberParagraph.Alignment = Element.ALIGN_CENTER;
	_doc.Add(vnNumberParagraph);
	
						Paragraph VNClosingParagraph = new Paragraph("Closing date: " + parsedDate.ToString("dd-MMMM-yyyy", CultureInfo.InvariantCulture), FontFactory.GetFont("Arial", 16, Font.BOLD));
						VNClosingParagraph.Alignment = Element.ALIGN_CENTER;
						_doc.Add(VNClosingParagraph);
						_doc.Add(new Chunk(Environment.NewLine));
						
						if(candsCount>0)
						{
							Paragraph candsCountParagraph = new Paragraph("Applicant Count: " + candsCount.ToString(), FontFactory.GetFont("Arial", 13, Font.BOLD));
							candsCountParagraph.Alignment = Element.ALIGN_CENTER;
							_doc.Add(candsCountParagraph);
							_doc.Add(new Chunk(Environment.NewLine));
						}
					}
					else
					{
						string firstPageTitle=VN + Environment.NewLine + DateTime.Now.ToString("dd-MMMM-yyyy", CultureInfo.InvariantCulture);
						if(!multi)
						{
							firstPageTitle = Org + " Individual Applicant CV" + Environment.NewLine + ApplicantName + Environment.NewLine + DateTime.Now.ToString("dd-MMMM-yyyy", CultureInfo.InvariantCulture);
						}
						Paragraph dateParagraph = new Paragraph(firstPageTitle, FontFactory.GetFont("Arial", 16, Font.BOLD));
						dateParagraph.Alignment = Element.ALIGN_CENTER;
						if(multi)
						{
							_doc.Add(dateParagraph);
							_doc.Add(new Chunk(Environment.NewLine));							
						}
						else if(InfoPage)
						{						
							_doc.Add(dateParagraph);
							_doc.Add(new Chunk(Environment.NewLine));
						}
						
						if(candsCount>0)
						{
							Paragraph candsCountParagraph = new Paragraph("Applicant Count: " + candsCount.ToString(), FontFactory.GetFont("Arial", 13, Font.BOLD));
							candsCountParagraph.Alignment = Element.ALIGN_CENTER;
							_doc.Add(candsCountParagraph);
							_doc.Add(new Chunk(Environment.NewLine));
						}
					}		
					
					_doc.NewPage();
					
					if(UseToc == "1" && cp != 0 && multi)
						AddPageWithInternalLinks(multi);				 

					_contentsTable = new List<TableOfContentsEntry>();
					_contentsTable1 = new List<TableOfContentsEntry>();
					
					
					if(!multi && !InfoPage)
					{
						try{
							iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(new Uri(urlImgLogo));
							jpg.ScaleToFit(50f, 50f);
							jpg.Alignment = iTextSharp.text.Image.TEXTWRAP ;
							jpg.SetAbsolutePosition(30f,_doc.PageSize.Height - 70f);
							_doc.Add(jpg);
						}
						catch{}													
					}
					
					
					//TOC
					Chapter chapter1 = new Chapter(new Paragraph("CV", FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL)), 0);
					chapter1.NumberDepth = 0;
					Section section1 = null;
					Paragraph paragraph3 = new Paragraph("", FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD));
					section1 = chapter1.AddSection(20f, paragraph3, 1);
					section1.NumberDepth = 0;
					section1.BookmarkOpen = false;
					section1.BookmarkTitle = "CV";
					_doc.Add(section1);	
					
					Section subsection1 = null;
			
					foreach (IElement element in elements)
					{
						var flag = true;
						
						String Test2= String.Empty;
						String Test3 = String.Empty;
						foreach (var item in element.Chunks)
						{
							//foreach (var item1 in item.Attributes)
							//{
							//    string key =item1.Key;
							//    string value =Convert.ToString(item1.Value);
							//}
							String content = item.Content.Trim();
							if (content.IndexOf("CAN") !=-1)
							{
								flag = false;
							}
							if(flag == false&&content != "CAN")
							{
								Test = content;
							}
							
							Test3 = Test2 = ReturnMatch(content);
						}				

						if (flag == false && Test != "CAN")
						{
							GlobalValue = Test;
							IsAppNameNeed = false;
						   if (Cont > 0)
						   {
								_doc.NewPage();
						   }
							Cont = Cont + 1;                                     

							Paragraph paragraph = new Paragraph();
							Anchor anchor1 = new Anchor(Test , FontFactory.GetFont("Arial", 13, Font.BOLD));
							//anchor1.Reference = "http://itextsharp.sourceforge.net";
							anchor1.Name = Test;
							paragraph.SpacingAfter = -11;
							paragraph.Alignment = Element.ALIGN_CENTER;
							paragraph.Add(anchor1);
							//_doc.Add(paragraph);
							
							subsection1 = chapter1.AddSection(10f, paragraph, 1);
							subsection1.BookmarkOpen = false;
							subsection1.BookmarkTitle = Test + ((tocNumbers && candsCount > 0) ? "  " + Cont + "/" + candsCount : "");
							subsection1.NumberDepth = 0; //NumberDepth=1 before
							_doc.Add(subsection1);
							
							string pagenumber = _pdfWriter.CurrentPageNumber.ToString();
							
							if(multi && Cont % 2 == 0)
								_contentsTable1.Add(new TableOfContentsEntry(Test,Test,pagenumber, ((tocNumbers && candsCount > 0) ? Cont + "/" + candsCount : "")));
							else
								_contentsTable.Add(new TableOfContentsEntry(Test,Test,pagenumber, ((tocNumbers && candsCount > 0) ? Cont + "/" + candsCount : "")));
						}
						else if (Test2 != String.Empty)
						{
							IsAppNameNeed = true;
							Paragraph paragraph = new Paragraph();
							Anchor anchor1;
							//clr == "#E9EBF3" for ILO
							if(clr == "#E9EBF3")
								anchor1 = new Anchor(Test3, FontFactory.GetFont("Arial", 10, Font.BOLD, BaseColor.BLACK));
							else						
								anchor1 = new Anchor(Test3, FontFactory.GetFont("Arial", 10, Font.BOLD, BaseColor.WHITE));
							//anchor1.Reference = "http://itextsharp.sourceforge.net";
							anchor1.Name = Test2 + Test3;						
							paragraph.Add(anchor1);
							PdfPTable table = new PdfPTable(4);
							table.HorizontalAlignment = 0;
							table.WidthPercentage = 100;
							PdfPCell cell = new PdfPCell(new Phrase(paragraph));
							cell.Colspan = 4;
							cell.HorizontalAlignment = 0; //0=Left, 1=Centre, 2=Right
							
							System.Drawing.Color col = System.Drawing.ColorTranslator.FromHtml(clr);
							BaseColor BS = new BaseColor(col);
							cell.BackgroundColor = BS; //it works
							table.AddCell(cell);
							//table.SpacingBefore = 1f;
							//table.SpacingAfter = 5f;
							
							_doc.Add(table);
							
							paragraph = new Paragraph();
							paragraph.SpacingAfter = -10;
							paragraph.Add(new Anchor(Test3, FontFactory.GetFont("Arial", 3, Font.BOLD, BaseColor.WHITE)));
							Section subSubSection = subsection1.AddSection(10f, paragraph, 2);
							subSubSection.BookmarkOpen = false;
							subSubSection.BookmarkTitle = Test3;
							subSubSection.NumberDepth = 0; //NumberDepth=2 before
							_doc.Add(subSubSection);

							string pagenumber = _pdfWriter.CurrentPageNumber.ToString();
							
							if(multi && Cont % 2 == 0)
								_contentsTable1.Add(new TableOfContentsEntry(Test2+Test3, Test3, pagenumber ));
							else
								_contentsTable.Add(new TableOfContentsEntry(Test2+Test3, Test3, pagenumber));	 
						}
						else
						{
							if (element is PdfPTable) 
							{
								PdfPTable tempEl = ((PdfPTable)element);
								tempEl.SplitLate = false;
								_doc.Add(tempEl);
							}
							else
								_doc.Add(element as IElement);
						}				
					}
					
					if(cp == 0)
						TotalPages = _pdfWriter.ReorderPages(null) + 1;
					
					if (UseToc == "1" && multi && cp == 0)
					{
					   AddPageWithInternalLinks(multi);
					}
					_doc.Close();				
					//Response.Redirect("Admin/Downloadpdf.asp?file=" + pdffilename); 
				}

				WebClient client = new WebClient();
				Byte[] buffer = client.DownloadData(dest_dir + pdffilename);
				//buffer = AddPageNumbers(buffer, viewFormat);
				
				if (buffer != null)
				{
					Response.ContentType = "application/pdf";
					Response.AddHeader("content-length", buffer.Length.ToString());
					Response.AddHeader("content-disposition", "attachment;filename=\"" + pdffilename + "\"");
					Response.BinaryWrite(buffer);
					Response.Flush();
					Response.End();
				}
			}
			catch (System.Threading.ThreadAbortException)
			{
				// ignore it
			}
			catch (Exception ex)
			{
				ErrorHendler(ex, Request.Url.ToString());	
				Response.Redirect("/err/er500.asp");
				
				//Response.Write(ex);	
			}
			
		}
		// For Individual CV Pdf  output
		else if (Request.QueryString["OutputType"] != null && Request.QueryString["OutputType"] == "PublicPdf")
		{			
			try
			{				  
				#region Get Html File From Query String

				String htmlfile = String.Empty;
				String clr = String.Empty;
				string pdfFileName = "";
				string ApplicantName = "";
				string A4Size= string.Empty;
				MinIndexPage = int.MinValue;
				MaxIndexpage = 1;
				int usepix = 0;
				bool multi = false;
				if (Request.QueryString["htmlfile"] != null)
				{
					htmlfile = (Request.QueryString["htmlfile"])+ ".html";
					if(htmlfile.Contains("..\\"))
					{
						htmlfile = htmlfile.Replace("..\\", "");
					}
					else if(htmlfile.Contains("\\"))
					{
						htmlfile = htmlfile.Replace("\\", "");
					}
					
					if(Request.QueryString["clr"]!=null&&Request.QueryString["clr"].ToString()!="")
						clr = "#"+(Request.QueryString["clr"]);
					else
						clr="#4000FF";
					VN = (Request.QueryString["Vnname"]);
					UseToc = Request.QueryString["UseToc"];
					A4Size = Request.QueryString["ASize"];
					VNClosing = "";
					Org = Request.QueryString["Org"];
					usepix = String.IsNullOrEmpty(Request.QueryString["Usepix"]) ? 0 : Convert.ToInt32(Request.QueryString["Usepix"]);
					viewFormat = String.IsNullOrEmpty(Request.QueryString["ViewFormat"]) ? 1 : Convert.ToInt32(Request.QueryString["ViewFormat"]);
					InfoPage = String.IsNullOrEmpty(Request.QueryString["InfoPage"]) ? InfoPage : (Request.QueryString["InfoPage"] == "Yes" ? true : false);
					TEXTSIZE = String.IsNullOrEmpty(Request.QueryString["TextSize"]) ? TEXTSIZE : Convert.ToSingle(Request.QueryString["TextSize"]);
					//if(!String.IsNullOrEmpty(Request.QueryString["PdfPath"]))
					//	PdfPath = Request.QueryString["PdfPath"];
					if(!String.IsNullOrEmpty(Request.QueryString["ApplicantName"]))
						ApplicantName = Regex.Replace(Request.QueryString["ApplicantName"], "[^a-zA-Z]", "");
				}
				#endregion 
				
				pdfFileName = "CV_" + ApplicantName + Path.GetFileNameWithoutExtension(htmlfile).Replace("CV_", "_") + ".pdf";
				String PdfFullPath = dest_dir + pdfFileName;
				
				// Step 8--One more requirement of ours was to process some input from certain web forms that might contain HTML tags and add them to generated .PDF, keeping the format intact. To accomplish that, we used iTextSharp.text.html.simpleparser.HTMLWorker as follows:
				//string stringWithHTMLTags = File.ReadAllText("E:\\docs\\UNSHARE\\short\\MULTI_SRCH_UNAIDS-LAWRENCER_20147113734.html", Encoding.Default);
				string stringWithHTMLTags = "";
				
				try	
				{
					stringWithHTMLTags = File.ReadAllText(dest_dir + htmlfile,Encoding.Default);
				}
				catch(DirectoryNotFoundException ex)
				{
					ErrorHendler(ex, Request.Url.ToString());
					
					Response.Write("File does not exist - level 1<br>");
					Response.Flush();
					Response.End();
				}				
				catch(FileNotFoundException ex)
				{
					//ErrorHendler(ex, Request.Url.ToString());
					
					Response.Write("File does not exist - level 2<br>");
					Response.Flush();
					Response.End();
				}
				
				
				if(usepix == 1)
					stringWithHTMLTags = stringWithHTMLTags.Replace("<img src=\"Doc", "<img src=\"" + imageFolder + "Doc");


				List<IElement> elements = new List<IElement>();
				try	
				{
					elements = ProcessHTML(stringWithHTMLTags);
				}
				catch(Exception ex)				
				{
					ErrorHendler(ex, Request.Url.ToString());
					
					string errormsg = "<br><br><br><div align='center' ><h1 style='color:red;'>Error occurred while generating the PDF</h1><br><table width='90%' >"+
									"<tr><td valign='top' bgcolor='#FFFBDF' colspan=2><font>This item has been logged with the Support Team and will be resolved soon.<br><Br>Should the error result in delay in your application, it will be noted and considered."+
									"<br>Otherwise, please allow 12 hours for non-essential issues to be resolved.<br></font></td></tr></table></div>";
					
					elements = ProcessHTML(errormsg);			
					
				}
				//we have to repeate because we want to keep internal links and bookmarks, also to have right numbering for pages and TOC
				for(int cp = 0; cp < repeateCount; cp++)
				{
					if (A4Size == "A4")
					{
						_doc = new Document(PageSize.A4, 30, 30, 30, 15);					
					}
					else if (A4Size == "Letter")
					{
						_doc = new Document(PageSize.LETTER, 30, 30, 30, 15);
					}
					else
					{
						_doc = new Document();
					}	
					

					_pdfWriter = PdfWriter.GetInstance(_doc, new FileStream(PdfFullPath, FileMode.Create));
					var PageMode = PdfWriter.PageLayoutTwoColumnLeft;
					if(viewFormat == 0)
						PageMode = PdfWriter.PageLayoutOneColumn;
					_pdfWriter.ViewerPreferences = PageMode;
					//_pdfWriter.ViewerPreferences = PdfWriter.PageModeUseOutlines;

					//Step2 After that make certain you set linear page mode, which ensures that page tree has a linear structure and allows you to insert pages in the middle of the document and  re-order them.
					_pdfWriter.SetLinearPageMode();
					 OpenDocument();
					 String  Test = String.Empty;
				
					 int Cont = 0;
					 _pdfWriter.PageEvent = new PageEventHelper();
					 _doc.NewPage();	
					 
					if(InfoPage)
					{
						Paragraph First = new Paragraph();
						First.Add(new Chunk("\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n"));
						_doc.Add(First);
						
						string firstPageTitle = Org + " Individual Applicant CV" + Environment.NewLine + ApplicantName + Environment.NewLine + DateTime.Now.ToString("dd-MMMM-yyyy", CultureInfo.InvariantCulture);;
						Paragraph dateParagraph = new Paragraph(firstPageTitle, FontFactory.GetFont("Arial", 20, Font.BOLD));
						dateParagraph.Alignment = Element.ALIGN_CENTER;
						
						_doc.Add(dateParagraph);
						_doc.Add(new Chunk(Environment.NewLine));
						_doc.NewPage();						
					}
					
					_contentsTable = new List<TableOfContentsEntry>();
				   
					//TOC
					Chapter chapter1 = new Chapter(new Paragraph("CV", FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL)), 0);
					chapter1.NumberDepth = 0;
					Section section1 = null;
					Paragraph paragraph3 = new Paragraph("", FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD));
					section1 = chapter1.AddSection(20f, paragraph3, 1);
					section1.NumberDepth = 0;
					section1.BookmarkOpen = false;
					section1.BookmarkTitle = "CV";
					_doc.Add(section1);			
					
					Section subsection1 = null;
						
					foreach (IElement element in elements)
					{
						var flag = true;					
						String Test2= String.Empty;
						String Test3 = String.Empty;
						foreach (var item in element.Chunks)
						{
							String content = item.Content.Trim();
							
							Test = content;							
							Test3 = Test2 = ReturnMatch(Test);
						}				

						if (Test2 != String.Empty)
						{
							Paragraph paragraph = new Paragraph();
							Anchor anchor1;
							//clr == "#E9EBF3" for ILO
							if(clr == "#E9EBF3")
								anchor1 = new Anchor(Test3, FontFactory.GetFont("Arial", 10, Font.BOLD, BaseColor.BLACK));
							else						
								anchor1 = new Anchor(Test3, FontFactory.GetFont("Arial", 10, Font.BOLD, BaseColor.WHITE));
				
							anchor1.Name = Test2 + Test3;
							paragraph.Add(anchor1);
							PdfPTable table = new PdfPTable(4);
							table.HorizontalAlignment = 0;
							table.WidthPercentage = 100;
							PdfPCell cell = new PdfPCell(new Phrase(paragraph));
							cell.Colspan = 4;
							cell.HorizontalAlignment = 0; //0=Left, 1=Centre, 2=Right
							
							System.Drawing.Color col = System.Drawing.ColorTranslator.FromHtml(clr);
							BaseColor BS = new BaseColor(col);
							cell.BackgroundColor = BS; //it works
							table.AddCell(cell);
							//table.SpacingBefore = 1f;
							//table.SpacingAfter = 5f;
							
							_doc.Add(table);
							
							paragraph = new Paragraph();
							paragraph.SpacingAfter = -10;
							paragraph.Add(new Anchor(Test3, FontFactory.GetFont("Arial", 3, Font.BOLD, BaseColor.WHITE)));
							subsection1 = chapter1.AddSection(10f, paragraph, 1);
							subsection1.BookmarkOpen = false;
							subsection1.BookmarkTitle = Test;
							subsection1.NumberDepth = 0;
							_doc.Add(subsection1);

							string pagenumber = _pdfWriter.CurrentPageNumber.ToString();
							
							_contentsTable.Add(new TableOfContentsEntry(Test2+Test3, Test3, pagenumber));	 
						}
						else
						{
							if (element is PdfPTable) 
							{
								PdfPTable tempEl = ((PdfPTable)element);
								tempEl.SplitLate = false;
								_doc.Add(tempEl);
							}
							else
								_doc.Add(element as IElement);
						}			
					}
					
					if(cp == 0)
						TotalPages = _pdfWriter.ReorderPages(null) + 1;
					
					if (UseToc == "1")
					{			
					   //AddPageWithInternalLinks(false);
					}
					_doc.Close();
					//Response.Redirect("public/pdf/public-view-f.asp?name=" + pdfFileName); 	
				}
				
				WebClient client = new WebClient();
				Byte[] buffer = client.DownloadData(PdfFullPath);
				if (buffer != null)
				{
					Response.ContentType = "application/pdf";
					Response.AddHeader("content-length", buffer.Length.ToString());
					Response.AddHeader("content-disposition", "attachment;filename=\"" + pdfFileName + "\"");
					Response.BinaryWrite(buffer);
					Response.Flush();
					Response.End();
				}
			}
			catch (System.Threading.ThreadAbortException)
			{
				// ignore it
			}
			catch (Exception ex)
			{
				ErrorHendler(ex, Request.Url.ToString());	
				Response.Redirect("/err/er500.asp");
				
				//Response.Write(ex);	
			}			
		}
		//For Public Pdf output
		else 
		{
			 try
			{				
				string Pdfname = String.Empty;
				string htmlfile = String.Empty;
				if (Request.QueryString["htmlfile"] != null)
				{
					htmlfile = (Request.QueryString["htmlfile"]) + ".html";
					if(htmlfile.Contains("..\\"))
					{
						htmlfile = htmlfile.Replace("..\\", "");
					}
					else if(htmlfile.Contains("\\"))
					{
						htmlfile = htmlfile.Replace("\\", "");
					}
				}
				
				if (Request.QueryString["PdfName"] != null&&Request.QueryString["PdfName"].Trim() != "")
				{
					Pdfname = Request.QueryString["PdfName"].Replace("{0}", DateTime.Now.ToString("yyyyMMdd")) + ".pdf";
				}
				else 
				{
					if (Request.QueryString["VnName"] != null&&Request.QueryString["VnName"].Trim() != "")
						Pdfname = Request.QueryString["VnName"] + ".pdf";
					else
						Pdfname = DateTime.Now.ToString("yyyyMMdd") + ".pdf";
				}
				
				bool arabic = false;
				string lng = Pdfname.Substring(Pdfname.Length - 6, 2);
				if(lng == "ar")	arabic = true;
				

				_doc = new Document(PageSize.A4, 10, 5, 10, 10);
				
				dest_dir = "E:\\docs\\UNSHARE\\PDFmake\\PDF-VN\\" + htmlfile;//.Replace("&" ,"-");
				if(Request.Url.Host.ToLower().Contains("wto"))
				{
					dest_dir = "E:\\docs\\WTO\\PDFmake\\PDF-VN\\" + htmlfile;
				}
				if(Request.Url.ToString().ToLower().Contains("/demo/"))
				{
					dest_dir = "E:\\docs\\stage\\ALLORG\\PDFmake\\PDF-VN\\" + htmlfile;//.Replace("&" ,"-");
				}
								
				
				//String PDf_dest_dir = "E:\\docs\\UNSHARE\\short\\";
				string[] words = dest_dir.Split('.');
				string pdfFileName = words[0] + "_" + UniqueID + ".pdf";
				
				// 25 SEP 23 LJL added TLS protocol to fix PDF output problem for VNs that occurred for days.
				System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

				_pdfWriter = PdfWriter.GetInstance(_doc, new FileStream(pdfFileName, FileMode.Create));
				OpenDocument();
				//var PageMode = PdfWriter.PageLayoutTwoColumnLeft;
				//_pdfWriter.ViewerPreferences = PageMode;
				_doc.NewPage();
	
				string stringWithHTMLTags = "";
				try	
				{
					stringWithHTMLTags = File.ReadAllText(dest_dir, Encoding.Default);		
				}
				catch(DirectoryNotFoundException ex)
				{
					ErrorHendler(ex, Request.Url.ToString());
					
					Response.Write("File does not exist - level 3<br>");
					Response.Flush();
					Response.End();
				}
				catch(FileNotFoundException ex)
				{
					//ErrorHendler(ex, Request.Url.ToString()); 
					
					Response.Write("File does not exist - level 4<br>");
					Response.Flush();
					Response.End();
				}
				
				List<IElement> elements = new List<IElement>();
				try	
				{
					elements = ProcessHTML(stringWithHTMLTags);
				}
				catch(Exception ex)				
				{
					ErrorHendler(ex, Request.Url.ToString());
					
					string errormsg = "<br><br><br><div align='center' ><h1 style='color:red;'>Error occurred while generating the PDF</h1><br><table width='90%' >"+
									"<tr><td valign='top' bgcolor='#FFFBDF' colspan=2><font>This item has been logged with the Support Team and will be resolved soon.<br><Br>Should the error result in delay in your application, it will be noted and considered."+
									"<br>Otherwise, please allow 12 hours for non-essential issues to be resolved.<br></font></td></tr></table></div>";
					
					elements = ProcessHTML(errormsg);			
					
				}

				foreach (var element in elements) 
				{
					if (element is PdfPTable) 
					{
						PdfPTable tempEl = ((PdfPTable)element);
						tempEl.SplitLate = false;
						if(arabic)
							tempEl.RunDirection = PdfWriter.RUN_DIRECTION_RTL;
							
						_doc.Add(tempEl);
					}
					else if ((element is PdfPCell) && arabic) 
					{
						PdfPCell celltmp = ((PdfPCell)element);
						celltmp.RunDirection = PdfWriter.RUN_DIRECTION_RTL;
						celltmp.HorizontalAlignment = Element.ALIGN_CENTER; 
						_doc.Add(celltmp);	 							
					}
					else
						_doc.Add(element as IElement);
				}

				CloseDocument();

				WebClient client = new WebClient();
				Byte[] buffer = client.DownloadData(pdfFileName);
				if (buffer != null)
				{
					Response.ContentType = "application/pdf";
					Response.AddHeader("content-length", buffer.Length.ToString());
					Response.AddHeader("content-disposition", "attachment;filename=\"" + Pdfname + "\"");
					Response.BinaryWrite(buffer);
					Response.Flush();
					Response.End();
				}
			}
			catch (System.Threading.ThreadAbortException)
			{
				// ignore it
			}
			catch (Exception ex)
			{					
				ErrorHendler(ex, Request.Url.ToString());	
				Response.Redirect("/err/er500.asp");
				
				//Response.Write(ex);				
			}
			finally 
			{	
				//_doc.CloseDocument();
			} 			
		}
    }


	private string ReturnMatch(string content)
	{
		if (content == "Personal Details")
			return content;
		else if (content == "Contract Information")
			return content;
		else if (content == "Language Skills")
			return content;
		else if (content == "Education")
			return content;
		else if (content == "International Experience")
			return content;
		else if (content == "Present and previous employment")
			return content;
		else if (content == "Computer Skills")
			return content;					
		else if (content == "Family Information" ||  content == "Family Info")
			return content;
		else if (content == "References")
			return content;
		else if (content == "Additional Information")
			return content;
		else if (content == "Clerical Skills")
			return content;
		else if (content == "Secretarial Skills")
			return content;
			
		else if (content == "Areas of Expertise")
			return content;  
		else if (content == "Other Information")
			return content;  
			
// 09 OCT 24 LJL added diversity to see if it works
		else if (content == "Diversity")
			return content;  

		else if (content == "Verification of Application")
			return content;  							
		else if (content == "Verification of Date and Place of Application to Vacancy")
			return content;
		else if (content == "Verification of Application to Vacancy")
			return content;
		else if (content.Contains("Covering Letter"))
			return content; 
		else if (content.Contains("Rotation and Mobility"))
			return content; 
		else if (content.Contains("Text Files -"))
			return content;
		else if (content.Contains("Additional Documents"))
			return content;
		
		
		
		content = content.Replace("\u00C9", "E"); 
		content = content.Replace("\u00E9", "e"); 
		content = content.Replace("\u00E0", "a"); 
		if (content == "Renseignements personnels")
			return content;
		else if (content == "Domaines de competence")
			return content;
		else if (content == "Competences linguistiques")
			return content;
		else if (content == "Etudes")
			return content;
		else if (content == "Renseignements familiaux")
			return content;
		else if (content.Contains("Emploi actuel et emplois precedents"))
			return content;
		else if (content == "Competences informatiques")
			return content;
		else if (content == "Renseignements complementaires")
			return content;
		else if (content.Contains("Lettre de couverture pour un avis de vacance specifique"))
			return content;
		else if (content == "Competences de secretariat")
			return content;
		else if (content.Contains("Verification de la date et du lieu de candidature a la vacance"))
			return content;							
		else if (content.Contains("Documents complementaires"))
			return content;	
		else if (content.Contains("Connaissances linguistiques"))
			return content;				
		else if (content.Contains("Formation"))
			return content;	
		else if (content.Contains("Types de contrats"))
			return content;		
		else if (content.Contains("Domaines de specialisation"))
			return content;	
		else if (content.Contains("Experience internationale"))
			return content;		
		else if (content.Contains("Poste actuel / precedent"))
			return content;		
		else if (content.Contains("Connaissances en informatique"))
			return content;		

// 09 OCT 24 LJL added diversity to see if it works
		else if (content.Contains("Diversité"))
			return content;		

			
		else if (content.Contains("References"))
			return content;		
		else if (content.Contains("Informations complementaires"))
			return content;
		return "";
		
	}
	  
	  
    /// <summary>
    /// Add a blank page to the document.
    /// </summary>
    /// <param name="doc"></param>
    private void AddPageWithInternalLinks(bool multi)
    {    
		IsAppNameNeed = false;
        _pageCount++;
		int totalpagesBeforeTOC=_pdfWriter.ReorderPages(null);
		
        PdfPTable _pdfContentsTable = new PdfPTable(2);        
        int totalpages = TotalPages;// _pdfWriter.ReorderPages(null)+2;
       
       
        MinIndexPage = _pdfWriter.PageNumber;      

        Anchor anchor2;
		PdfPTable table = new PdfPTable(2);
		table.WidthPercentage = 100;
		if(multi)
		{
			String Head = "";
			for (int i = 0; i < _contentsTable.Count; i++) 
			{
				var content = _contentsTable[i];
				
				if(content.Title.Contains("Covering Letter for a specific vacancy"))
				{
					content.Title = "Covering Letter for a specific vacancy";
				}
				
				if(content.Title.Contains("Text Files"))
				{
					content.Title = "Text Files";
				}
				
				content.Page = (Convert.ToInt32(content.Page) + TocPagesCount).ToString();	
				if(ShowTocPageNumbers)		
					Head = "      {0} {1}" + content.Page + "/" + totalpages;
				else
					Head = "      {0} {1}";
				
				switch (content.Title)
				{
					case "Personal Details":
						Head = Head.Replace("{0}", "Personal Details").Replace("{1}", AddDots(80));
						break;
					case "Contract Information":
						Head = Head.Replace("{0}", "Contract Information").Replace("{1}", AddDots(74));
						break;
					case "Areas of Expertise":
						Head = Head.Replace("{0}", "Areas of Expertise").Replace("{1}", AddDots(78));
						break;
					case "Language Skills":
						Head = Head.Replace("{0}", "Language Skills").Replace("{1}", AddDots(81));
						break;
					case "Education":
						Head = Head.Replace("{0}", "Education").Replace("{1}", AddDots(91));
						break;
					case "International Experience":
						Head = Head.Replace("{0}", "International Experience").Replace("{1}", AddDots(68));
						break;
					case "Present and previous employment":
						Head = Head.Replace("{0}", "Present and previous employment").Replace("{1}", AddDots(52));
						break;
					case "Family Information":
						Head = Head.Replace("{0}", "Family Information").Replace("{1}", AddDots(77));
						break;
					case "Family Info":
						Head = Head.Replace("{0}", "Family Information").Replace("{1}", AddDots(77));
						break;
					case "References":
						Head = Head.Replace("{0}", "References").Replace("{1}", AddDots(89));
						break;
					case "Additional Information":
						Head = Head.Replace("{0}", "Additional Information").Replace("{1}", AddDots(72));
						break;
					case "Other Information":
						Head = Head.Replace("{0}", "Other Information").Replace("{1}", AddDots(79));
						break;
					case "Clerical Skills":
						Head = Head.Replace("{0}", "Clerical Skills").Replace("{1}", AddDots(86));
						break;

// 09 OCT 24 LJL added diversity to see if it works
					case "Diversity":
						Head = Head.Replace("{0}", "Diversity").Replace("{1}", AddDots(93));
						break;


					case "Secretarial Skills":
						Head = Head.Replace("{0}", "Secretarial Skills").Replace("{1}", AddDots(81));
						break;
					case "Computer Skills":
						Head = Head.Replace("{0}", "Computer Skills").Replace("{1}", AddDots(82));
						break;
					case "Verification of Date and Place of Application to Vacancy":
						Head = Head.Replace("{0}", "Verification of Date and Place of Application to Vacancy").Replace("{1}", AddDots(18));
						break;
					case "Verification of Application to Vacancy":
						Head = Head.Replace("{0}", "Verification of Application to Vacancy").Replace("{1}", AddDots(47));
						break;
					case "Covering Letter for a specific vacancy":
						Head = Head.Replace("{0}", "Covering Letter for a specific vacancy").Replace("{1}", AddDots(47));
						break;
					case "Rotation and Mobility":
						Head = Head.Replace("{0}", "Rotation and Mobility").Replace("{1}", AddDots(73));
						break;	
					case "Verification of Application":
						Head = Head.Replace("{0}", "Verification of Application").Replace("{1}", AddDots(66));
						break;	
					case "Text Files":
						Head = Head.Replace("{0}", "Text Files").Replace("{1}", AddDots(91));
						break;
					default:
					{
						if(ShowTocPageNumbers)		
							Head = content.Title + " (" + content.ApplicantsCount + ")"; 
						else
							Head = content.Title; 
						
						break;
					}					
				}				

				anchor2 = new Anchor(Head, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL));                
				anchor2.Reference = "#" + content.TitleFull;
				Paragraph paragraph2 = new Paragraph("", FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL));
				paragraph2.Add(anchor2);
								
				
				PdfPCell cell = new PdfPCell(new Phrase(paragraph2));
				cell.Border  = 0;
				//cell.Colspan = 4;
				cell.HorizontalAlignment = 0; //0=Left, 1=Centre, 2=Right
				
				table.AddCell(cell);
				
				//tmpTitle = "";
				
				if(_contentsTable1.Count > i)
				{
					var content1 = _contentsTable1[i];
					
					if(content1.Title.Contains("Covering Letter for a specific vacancy"))
					{
						content1.Title = "Covering Letter for a specific vacancy";
					}	
					if(content1.Title.Contains("Text Files"))
					{
						content1.Title = "Text Files";
					}					
		
					content1.Page = (Convert.ToInt32(content1.Page) + TocPagesCount).ToString();	
				
					if(ShowTocPageNumbers)		
						Head = "      {0} {1}" + content1.Page + "/" + totalpages;
					else
						Head = "      {0} {1}";
					switch (content1.Title)
					{
						case "Personal Details":
							Head = Head.Replace("{0}", "Personal Details").Replace("{1}", AddDots(80));
							break;
						case "Contract Information":
							Head = Head.Replace("{0}", "Contract Information").Replace("{1}", AddDots(74));
							break;
						case "Areas of Expertise":
							Head = Head.Replace("{0}", "Areas of Expertise").Replace("{1}", AddDots(77));
							break;
						case "Language Skills":
							Head = Head.Replace("{0}", "Language Skills").Replace("{1}", AddDots(81));
							break;
						case "Education":
							Head = Head.Replace("{0}", "Education").Replace("{1}", AddDots(91));
							break;
						case "International Experience":
							Head = Head.Replace("{0}", "International Experience").Replace("{1}", AddDots(67));
							break;
						case "Present and previous employment":
							Head = Head.Replace("{0}", "Present and previous employment").Replace("{1}", AddDots(52));
							break;
						case "Family Information":
							Head = Head.Replace("{0}", "Family Information").Replace("{1}", AddDots(77));
							break;
						case "Family Info":
							Head = Head.Replace("{0}", "Family Information").Replace("{1}", AddDots(77));
							break;
						case "References":
							Head = Head.Replace("{0}", "References").Replace("{1}", AddDots(88));
							break;
						case "Additional Information":
							Head = Head.Replace("{0}", "Additional Information").Replace("{1}", AddDots(72));
							break;
						case "Other Information":
							Head = Head.Replace("{0}", "Other Information").Replace("{1}", AddDots(79));
							break;
						case "Clerical Skills":
							Head = Head.Replace("{0}", "Clerical Skills").Replace("{1}", AddDots(86));
							break;

// 09 OCT 24 LJL added diversity to see if it works
						case "Diversity":
							Head = Head.Replace("{0}", "Diversity").Replace("{1}", AddDots(93));
							break;


						case "Secretarial Skills":
							Head = Head.Replace("{0}", "Secretarial Skills").Replace("{1}", AddDots(81));
							break;
						case "Computer Skills":
							Head = Head.Replace("{0}", "Computer Skills").Replace("{1}", AddDots(81));
							break;
						case "Verification of Date and Place of Application to Vacancy":
							Head = Head.Replace("{0}", "Verification of Date and Place of Application to Vacancy").Replace("{1}", AddDots(18));
							break;
						case "Verification of Application to Vacancy":
							Head = Head.Replace("{0}", "Verification of Application to Vacancy").Replace("{1}", AddDots(47));
							break;
						case "Verification of Application":
							Head = Head.Replace("{0}", "Verification of Application").Replace("{1}", AddDots(66));
							break;	
						case "Covering Letter for a specific vacancy":
							Head = Head.Replace("{0}", "Covering Letter for a specific vacancy").Replace("{1}", AddDots(47));
							break;
						case "Rotation and Mobility":
							Head = Head.Replace("{0}", "Rotation and Mobility").Replace("{1}", AddDots(73));
							break;
						case "Text Files":
							Head = Head.Replace("{0}", "Text Files").Replace("{1}", AddDots(91));
							break;							
						default:
						{
							if(ShowTocPageNumbers)		
								Head = content1.Title + " (" + content1.ApplicantsCount + ")"; 
							else
								Head = content1.Title; 
								break;
						}						
					}				

					Anchor anchor3 = new Anchor(Head, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL));                
					anchor3.Reference = "#" + content1.TitleFull;
					Paragraph paragraph3 = new Paragraph("", FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL));
					paragraph3.Add(anchor3);
									
					
					PdfPCell cell1 = new PdfPCell(new Phrase(paragraph3));
					cell1.Border  = 0;
					//cell.Colspan = 4;
					cell1.HorizontalAlignment = 0; //0=Left, 1=Centre, 2=Right
					
					table.AddCell(cell1);
				}
				else
				{
					PdfPCell cell1 = new PdfPCell(new Phrase(""));
					cell1.Border  = 0;
					table.AddCell(cell1);
				}				
			}				
		}        
		else
		{
		}
		
		_doc.Add(table);
		_doc.NewPage();
       // _doc.Add(new Chunk(Environment.NewLine));
		
		if(_pdfWriter.ReorderPages(null)+1 > TotalPages)
		{
			TotalPages=_pdfWriter.ReorderPages(null)+1;
			TocPagesCount = TotalPages - 1 - totalpagesBeforeTOC;
		}
		
        /** Reorder pages so that TOC will will be the second page in the doc
           * right after the title page**/
        //MaxIndexpage = _pdfWriter.PageNumber;
      /*  int toc = _pdfWriter.PageNumber - 1;
        int total = _pdfWriter.ReorderPages(null);
        int tocnum=total-totalpagesBeforeTOC;
		TocPagesCount = tocnum;
        int[] order = new int[total];

        int toc1 = totalpagesBeforeTOC+1;       
       //_doc.Add(new Chunk(toc1.ToString() + " - " + totalpagesBeforeTOC + " " + total+ " " + toc));
        for (int i = 0; i < total; i++)
        {            
            if (i == 0)
            {
                order[i] = 1;
            }
            else if (i >= 1 && i <= tocnum)
            {                
                order[i] = toc1;
                toc1 = toc1 + 1;              
            }
            else //if(i <=totalpagesBeforeTOC+1)
            {
                order[i] = i - tocnum+1;
            }            
        }

        _pdfWriter.ReorderPages(order);	*/
    }
	
	private string AddDots(int count)
    {
		string dots = "";
		for (int i = 0; i < count; i++)
		{
			dots += ".";
		}
		return dots;
	}
    public static System.Drawing.SizeF MeasureString(string s, System.Drawing.Font font)
    {
        System.Drawing.SizeF result;
        using (var image = new System.Drawing.Bitmap(1, 1))
        {
            using (var g = System.Drawing.Graphics.FromImage(image))
            {
                result = g.MeasureString(s, font);
            }
        }

        return result;
    }

   
    /// <summary>
    /// Add a paragraph object containing the specified element to the PDF document.
    /// </summary>
    /// <param name="doc">Document to which to add the paragraph.</param>
    /// <param name="alignment">Alignment of the paragraph.</param>
    /// <param name="font">Font to assign to the paragraph.</param>
    /// <param name="content">Object that is the content of the paragraph.</param>
    private void AddParagraph(Document doc, int alignment, iTextSharp.text.Font font, iTextSharp.text.IElement content)
    {
        Paragraph paragraph = new Paragraph();
        paragraph.SetLeading(0f, 1.2f);
        paragraph.Alignment = alignment;
        paragraph.Font = font;
        paragraph.Add(content);
        doc.Add(paragraph);
    } 
    //Step15 --To be able to access generated .PDF document, I just created a custom property which exposes a byte array of the _output MemoryStream object I’ve created at the start.
    public Byte[] DocumentContents
    {
        get { return _output.ToArray(); }
    }

    //Step--14 And that’s it. Once you register it as a page event, iTextSharp triggers OnEndPage() at appropriate times and puts _docBaseName string (or whatever you want) at the bottom of every page.
    //Finally you have to close the document and you are done!
    public void CloseDocument()
    {
        _doc.Close();
        //_doc.Dispose();
    }
    //Step3 Second step is opening the document and after that you are pretty much ready to go.
    public void OpenDocument()
    {
        _doc.Open();
    }

    //Step8 
    private List<IElement> ProcessHTML(string strToProcess)
    {
		string arialuniTff = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "ARIAL.TTF");
		//Register the font with iTextSharp
		iTextSharp.text.FontFactory.Register(arialuniTff);
		
		string textSize = TEXTSIZE.ToString() + "pt";
        StyleSheet styles = new StyleSheet();
		styles.LoadTagStyle(HtmlTags.BODY, HtmlTags.FACE, "Arial Unicode MS");
		styles.LoadTagStyle(HtmlTags.BODY, HtmlTags.ENCODING, BaseFont.IDENTITY_H);

        styles.LoadTagStyle(HtmlTags.BODY, HtmlTags.SIZE, textSize);		//Just number also works 12pt=4
        styles.LoadTagStyle(HtmlTags.TD, HtmlTags.SIZE, textSize);
		//styles.LoadTagStyle(HtmlTags.BODY, HtmlTags.SIZE, "9px");		
        //styles.LoadTagStyle(HtmlTags.TD, HtmlTags.SIZE, "9px");
        styles.LoadTagStyle(HtmlTags.H2, HtmlTags.SIZE, "12px");
        styles.LoadTagStyle(HtmlTags.TABLE, HtmlTags.PADDINGLEFT, "2px");   

		//strToProcess = Server.HtmlEncode(strToProcess);     
		
		CleanHtmlString(ref strToProcess);	
        
		//strToProcess = "<html><head></head><body><br/><br/><h2 style='color:#FF0000'>Image testing</h2><img src='E:\\docs\\UNSHARE\\vac-cv\\Doc262873.JPG' alt='' width='150'></body></html>";
        return HTMLWorker.ParseToList(new StringReader(strToProcess), styles);        
    }
	
    private void CleanHtmlString(ref string strToProcess)
    {
		  string pattern = @"<([0-9])";
		  string replacement = "$1";
		  strToProcess = System.Text.RegularExpressions.Regex.Replace(strToProcess, pattern, replacement);       
    }	
	
	public byte[] AddPageNumbers(byte[] pdf, int viewFormat)
	{
		MemoryStream ms = new MemoryStream();
		ms.Write(pdf, 0, pdf.Length);
		
		PdfReader reader = new PdfReader(pdf);
		// we retrieve the total number of pages
		// we exclude the last page from the PDF because it is empty for multi PDF
		int n = reader.NumberOfPages - 1;
		// we retrieve the size of the first page
		Rectangle psize = reader.GetPageSize(1);

		var PageMode = PdfWriter.PageLayoutTwoColumnLeft;
		if(viewFormat == 0)
			PageMode = PdfWriter.PageLayoutOneColumn;
			
		Document document = new Document(psize,  30, 15, 25, 25);
		// step 2: we create a writer that listens to the document
		PdfWriter writer = PdfWriter.GetInstance(document, ms);
		writer.ViewerPreferences = PageMode;

		document.Open();
		// add content
		PdfContentByte cb = writer.DirectContent;

		int p = 0;
		for (int page = 1; page <= reader.NumberOfPages - 1; page++) //the last page is empty for multi
		{
			document.NewPage();
			p++;

			PdfImportedPage importedPage = writer.GetImportedPage(reader, page);
			cb.AddTemplate(importedPage, 0, 0);

			BaseFont bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
			cb.BeginText();
			BaseColor grey = new BaseColor(128, 128, 128);
		    cb.SetColorFill(grey);
			cb.SetFontAndSize(bf, 7);
			cb.ShowTextAligned(PdfContentByte.ALIGN_RIGHT, + p + "/" + n, 560, 20, 0);
			cb.EndText();
		}
		
		document.Close();
		return ms.ToArray();
	}
	
	/// <summary>
    /// Return a unique identifier based on system's full date (yyyymmdd) and time (hhmissms).
    /// 
    /// Output sample: 2006040212445099
    /// </summary>
    public static string UniqueID
    {
      get
      {
        DateTime date = DateTime.Now;

        string uniqueID = String.Format(
          "{0:0000}{1:00}{2:00}{3:00}{4:00}{5:00}{6:000}",
          date.Year, date.Month, date.Day,
          date.Hour, date.Minute, date.Second, date.Millisecond
          );
        return uniqueID;
      }
    }
	
	private void ErrorHendler(Exception ex, string url)
    {
		string Org = Request.QueryString["Org"];
		string OrgCode = Request.QueryString["OrgCode"];
		string errMsg = ex.InnerException != null ? ex.Message  + "\n" +  ex.InnerException.Message : ex.ToString().Replace("System.NullReferenceException: ","");
		
		try
		{			
			if(errMsg.IndexOf("line ") <= 0)
			{
				errMsg += " " + ex.StackTrace +" " + ex.Source;
			}
			
			var line = "0";
			int lineCharCount = errMsg.IndexOf(" ",(errMsg.IndexOf("line ")+ 5)) - (errMsg.IndexOf("line ")+ 5);
			if(lineCharCount <= 0)
			{
				lineCharCount = errMsg.Length - (errMsg.IndexOf("line ")+ 5);
			}
			
			try	
			{
				line = errMsg.IndexOf("line ") > 0 ? errMsg.Substring(errMsg.IndexOf("line ")+ 5, lineCharCount) : "0";
			
				errMsg = (errMsg.Length <= 500 ? errMsg : 
					errMsg.Substring(0, 200).Replace("'", "''") + 
					errMsg.Substring(errMsg.Length-300, errMsg.Length -(errMsg.Length-300)).Replace("'", "''"));
			}
			catch
			{}
			
			string connectionString = "Data Source=192.168.106.5;Initial Catalog=iislogs;User Id=write_human;Password=act1ve44";
			 
			string INS1 = "crd_d, serr_sitetype, serr_dsc_c, ";
			string INS2 = "getdate(), 'PDF-" + Org + "', '" + errMsg + "',";

			if (!String.IsNullOrEmpty(ex.Source)) 
			{
				INS1 += "serr_source_c, ";
				INS2 += " '" + (ex.Source.Length <= 500 ? ex.Source+ "Test" : ex.Source.Substring(0, 500).Replace("'", "''")) +  "',";
			}

			INS1 += "serr_file_c, ";
			INS2 += " '" + url +  "',";

			if (Session["CLI_RSYS_ADMIN_USER"] != null && !String.IsNullOrEmpty(Session["CLI_RSYS_ADMIN_USER"].ToString()))
			{
				INS1 = INS1 + "user_id_t, ";
				INS2 = INS2 + " '" + (Session["CLI_RSYS_ADMIN_USER"].ToString().Length <= 150 ? Session["CLI_RSYS_ADMIN_USER"].ToString() : Session["CLI_RSYS_ADMIN_USER"].ToString().Substring(0, 150))  +  "',";
			}
			else if (Session["RSYSUSER"] != null && !String.IsNullOrEmpty(Session["RSYSUSER"].ToString()))
			{
				INS1 += "user_id_t, ";
				INS2 += " '" + (Session["RSYSUSER"].ToString().Length <= 150 ? Session["RSYSUSER"].ToString() : Session["RSYSUSER"].ToString().Substring(0, 150))  +  "',";
			}

			INS1 += "serr_thisorg_c, ";
			INS2 += " '" + OrgCode +"',";
			

			
			INS1 += "serr_line_c, ";
			INS2 +=  " '" + line + "',";

			if (!String.IsNullOrEmpty(Request.ServerVariables["HTTP_USER_AGENT"].ToString()))
			{
				INS1 += "serr_user_browser_c, ";
				INS2 += " '" + (Request.ServerVariables["HTTP_USER_AGENT"].ToString().Length <= 120 	
						? Request.ServerVariables["HTTP_USER_AGENT"].ToString() 
						: Request.ServerVariables["HTTP_USER_AGENT"].ToString().Substring(0, 120))  +  "',";
			}
			if (Session["HTTP_REFERER"] != null && !String.IsNullOrEmpty(Request.ServerVariables["HTTP_REFERER"].ToString()))
			{
				INS1 += "serr_web_refer_c, ";
				INS2 += " '" + (Request.ServerVariables["HTTP_REFERER"].ToString().Length <= 120 	
						? Request.ServerVariables["HTTP_REFERER"].ToString() 
						: Request.ServerVariables["HTTP_REFERER"].ToString().Substring(0, 120))  +  "',";
			}
			if (Session["HTTP_REFERER"] != null && !String.IsNullOrEmpty(Session["HTTP_REFERER"].ToString()))
			{
				INS1 += "cand_id_c, ";
				INS2 += " '" + Session["RSYS_EVAL"].ToString() + "',";
			}


			if (Request.Form != null && Request.Form["jobinfo_uid_c"] != null && !String.IsNullOrEmpty(Request.Form["jobinfo_uid_c"].ToString()))
			{
				INS1 += "jobinfo_uid_c, ";
				INS2 += " '" + Request.Form["jobinfo_uid_c"]  + "',";
			}
			else if (Request.Form["jobinfo_uid_c"] != null && !String.IsNullOrEmpty(Request.Form["jafid"]))
			{
				INS1 += "jobinfo_uid_c, ";
				INS2 += " '" + Request.Form["jafid"]  + "',";
			}
			else if (Request.Form["jobinfo_uid_c"] != null && !String.IsNullOrEmpty(Request.Form["jobid"]))
			{
				INS1 += "jobinfo_uid_c, ";
				INS2 += " '" + Request.Form["jobid"] + "',";
			}
			else if (Request.Form["jobinfo_uid_c"] != null && !String.IsNullOrEmpty(Request.Form["jobid"]))
			{
				INS1 += "jobinfo_uid_c, ";
				INS2 += " '" + Request.Form["jobid"]  + "',";
			}

			string queryString =
				"INSERT INTO td_serr ( " + 
				INS1 + " serr_ip_c ) VALUES ( " +
				INS2 +
				" '" +
				Request.ServerVariables["REMOTE_ADDR"] + "') ";
				 

			using (SqlConnection connection = new SqlConnection(connectionString))
			{
				SqlCommand command = new SqlCommand(queryString, connection);
				//command.Parameters.AddWithValue("@pricePoint", paramValue);

			   {
					connection.Open();
					command.ExecuteNonQuery(); 
					connection.Close(); 
				}
				
			}
		}
		catch (Exception exm)
		{
		}
	}
}

//Step--10 If you need to create table of contents for your document, I suggest doing it at the end of document generation, right before you close it. That way, you already know the structure and are ready to create a table representing this structure. The way I approached this is first I’ve created a struct representing a TOC entry.
public struct TableOfContentsEntry
{
    private string _titleFull;
	private string _title;
    private string _page;
    private string _applicantsCount;	
	
    public TableOfContentsEntry(string titleFull,string title, string page)
    {
		_titleFull = titleFull;
        _title = title;
        _page = page;
		_applicantsCount = "";
    }  
	
    public TableOfContentsEntry(string titleFull,string title, string page, string applicantsCount)
    {
		_titleFull = titleFull;
        _title = title;
        _page = page;
		_applicantsCount = applicantsCount;
    }  
	
	public string TitleFull
    {
        get { return _titleFull; }
        set { _titleFull = value; }
    }

    public string Title
    {
        get { return _title; }
        set { _title = value; }
    }

    public string Page
    {
        get { return _page; }
        set { _page = value; }
    }

    public string ApplicantsCount
    {
        get { return _applicantsCount; }
        set { _applicantsCount = value; }
    }
}


public class PageEventHelper : PdfPageEventHelper
{   
    protected Font footer
    {
        get
        {
            // create a basecolor to use for the footer font, if needed.
            BaseColor grey = new BaseColor(128, 128, 128);
            Font font = FontFactory.GetFont("Arial", 7, Font.NORMAL, grey);
            return font;
        }
    }  

	public override void OnEndPage(PdfWriter writer, Document doc)
	{
		PdfPTable footerTbl = new PdfPTable(3);
		footerTbl.TotalWidth = doc.PageSize.Width;
		footerTbl.HorizontalAlignment = Element.ALIGN_CENTER;

		string formatted = DateTime.Now.ToString("dd-MMMM-yyyy  hh:mm:ss", CultureInfo.InvariantCulture);
	   
		Paragraph para = new Paragraph("Printed: " + formatted, footer);
		//add a carriage return
		//para.Add(Environment.NewLine);

		PdfPCell cell = new PdfPCell(para);
		cell.Border = 0;
		cell.PaddingLeft = 10;
		cell.PaddingTop = 8;
		footerTbl.AddCell(cell);	

		//create new instance of Paragraph for 2nd cell text
		if(!String.IsNullOrEmpty(GeneratePDF.VNClosing) && GeneratePDF.VNClosing.Trim() != "")
		{		
			CultureInfo enUS = new CultureInfo("en-US"); 
			DateTime parsedDate = DateTime.Now;
			try
			{
				parsedDate = DateTime.ParseExact(GeneratePDF.VNClosing, "d", enUS);
			}
			catch
			{
				//parsedDate = DateTime.ParseExact(GeneratePDF.VNClosing, "dd MM yy", enUS);
				try
				{
					parsedDate = DateTime.ParseExact(GeneratePDF.VNClosing, "dd MM yyyy", enUS);
				}
				catch
				{
					parsedDate = Convert.ToDateTime(GeneratePDF.VNClosing);
				}
			}		
			
			para = new Paragraph("Closing date: " + parsedDate.ToString("dd-MMMM-yyyy", CultureInfo.InvariantCulture), footer);
		}
		else
		{
			para = new Paragraph("", footer);//GeneratePDF.VN
		}
		cell = new PdfPCell(para);
		cell.HorizontalAlignment = Element.ALIGN_CENTER;
		cell.Border = 0;
		cell.PaddingRight = 10;
		cell.PaddingTop = 8;
		footerTbl.AddCell(cell);

		string pagenumber = String.Empty;
		/*if (GeneratePDF.MinIndexPage != int.MinValue)
		{
			pagenumber = GeneratePDF.NumberToRoman(GeneratePDF.MaxIndexpage);
			GeneratePDF.MaxIndexpage = GeneratePDF.MaxIndexpage + 1;
		}
		else*/
		{
		   pagenumber = Convert.ToString(doc.PageNumber);
		}
		
		para = new Paragraph("Page " + pagenumber + "/" + GeneratePDF.TotalPages.ToString(), footer);
		//para = new Paragraph("");
		cell = new PdfPCell(para);
		cell.HorizontalAlignment = Element.ALIGN_RIGHT;
		cell.Border = 0;
		cell.PaddingRight = 20;
		cell.PaddingTop = 8;
		footerTbl.AddCell(cell);

		//write the rows out to the PDF output stream.
		footerTbl.WriteSelectedRows(0, -1, 0, (doc.BottomMargin + 10), writer.DirectContent);
	}

	//override the OnStartPage event handler to add our header
	public override void OnStartPage(PdfWriter writer, Document doc)
	{
		Paragraph Header = new Paragraph(); 
		if(GeneratePDF.IsAppNameNeed)
			Header = new Paragraph(GeneratePDF.GlobalValue, footer);
		else
			Header = new Paragraph("", footer);
		
		Header.Alignment = Element.ALIGN_LEFT;

		//Paragraph Header1 = new Paragraph("", footer);
		Paragraph Header2 = new Paragraph(GeneratePDF.Org + "  " + GeneratePDF.VN, footer);//DateTime.Now.ToString("dd-MMMM-yy", CultureInfo.InvariantCulture)

		PdfPTable HeaderTbl = new PdfPTable(2);
		HeaderTbl.TotalWidth = doc.PageSize.Width;
		HeaderTbl.HorizontalAlignment = Element.ALIGN_RIGHT;
		
		float[] widths = new float[] { 1f, 3f };
		HeaderTbl.SetWidths(widths);

		PdfPCell cell = new PdfPCell(Header);
		cell.HorizontalAlignment = Element.ALIGN_LEFT;
		cell.Border = 0;
		cell.PaddingLeft  = 10;
		HeaderTbl.AddCell(cell);

		/*cell = new PdfPCell(Header1);
		cell.HorizontalAlignment = Element.ALIGN_CENTER;
		cell.Border = 0;
		cell.PaddingRight = 10;
		HeaderTbl.AddCell(cell);*/

		cell = new PdfPCell(Header2);
		cell.HorizontalAlignment = Element.ALIGN_RIGHT;
		cell.Border = 0;
		cell.PaddingRight = 10;
		HeaderTbl.AddCell(cell);
		
		HeaderTbl.WriteSelectedRows(0, -1, 0, (doc.PageSize.Height - 10), writer.DirectContent);		
	} 
}