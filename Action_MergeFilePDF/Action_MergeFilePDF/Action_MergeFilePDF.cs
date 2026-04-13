using Microsoft.Xrm.Sdk;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static PdfSharp.Pdf.PdfDictionary;
namespace Action_MergeFilePDF
{
    public class Action_MergeFilePDF : IPlugin
    {

        IPluginExecutionContext context = null;
        IOrganizationServiceFactory factory = null;
        IOrganizationService service = null;
        ITracingService tracingService = null;
        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            context = serviceProvider.GetService(typeof(IPluginExecutionContext)) as IPluginExecutionContext;
            factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            service = (IOrganizationService)factory.CreateOrganizationService(context.UserId);
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var files = context.InputParameters["files"].ToString().Split(',');
            tracingService.Trace("count: " + files.Count());
            var filers = MergePdfFiles(files);
            context.OutputParameters["fileres"] = filers;
        }
        public string MergePdfFiles(string[] base64Files)
        {
            tracingService.Trace("validBase64Files:" + base64Files[0]);
            var validBase64Files = base64Files.Where(f => !string.IsNullOrWhiteSpace(f)).ToList();
            if (validBase64Files.Count == 0)
            {
                tracingService.Trace("Không có file hợp lệ nào để gộp.");
                return string.Empty;
            }
            int count = 0;
            // --- LOGIC MỚI BẮT ĐẦU TỪ ĐÂY ---
            using (PdfDocument outputDocument = new PdfDocument())
            {


                // 2. Lặp qua TẤT CẢ các file đầu vào của bạn
                foreach (var base64File in validBase64Files)
                {
                    try
                    {
                        byte[] pdfBytes = Convert.FromBase64String(base64File);
                        using (MemoryStream pdfStream = new MemoryStream(pdfBytes))
                        // Mở TẤT CẢ các file ở chế độ Import an toàn
                        using (PdfDocument inputDocument = PdfReader.Open(pdfStream, PdfDocumentOpenMode.Import))
                        {
                            foreach (PdfPage page in inputDocument.Pages)
                            {
                                outputDocument.AddPage(page);
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace($"Lỗi khi import một file PDF: {ex.Message}");
                        throw new InvalidPluginExecutionException($"Một trong các file PDF không hợp lệ. Chi tiết: {ex.Message}", ex);
                    }
                    count++;
                }





                // 4. Lưu kết quả cuối cùng
                using (MemoryStream resultStream = new MemoryStream())
                {
                    outputDocument.Save(resultStream, false);
                    return Convert.ToBase64String(resultStream.ToArray());
                }
            }
        }
    }
}
