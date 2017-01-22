using System.IO;
using System.Reflection;
using Aspose.Words.Reporting;

namespace Aspose.Words.Samples
{
    public class Linq
    {
        private const string TemplateFileName = "SampleDocTemplate.docx";
        private const string OutputFileName = "SampleDoc.docx";

        public static void Run()
        {
            var folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            var templateFilePath = Path.Combine(folder, TemplateFileName);
            var outFilePath = Path.Combine(folder, OutputFileName);

            var template = new Document(templateFilePath);

            var document = SampleDoc.Create();

            var engine = new ReportingEngine();
            engine.BuildReport(template, document, "Doc");

            template.Save(outFilePath);
        }
    }
}