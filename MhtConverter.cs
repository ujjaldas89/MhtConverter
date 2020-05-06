using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HtmlAgilityPack;

namespace MhtConverter
{
    public class MhtConverter
    {
        private Dictionary<string, string> _scriptSources;
        private Dictionary<string, string> _externalResource;
        public string OutputFile { get; private set; }
        private string _sourceHtmlDirectory;
        private string _targetFile;
        private string _tempHtmlFile;
        private string _sourceHtmlFile;
        public MhtConverter()
        {
            _scriptSources = new Dictionary<string, string>();
            _externalResource = new Dictionary<string, string>();
        }

        public void Convert(string htmlFile, string targetFile = null)
        {
            _sourceHtmlFile = htmlFile;
            _targetFile = targetFile;
            _tempHtmlFile = GetTempHtml();
            _sourceHtmlDirectory = Path.GetDirectoryName(htmlFile);
            var htmlRead = new StreamReader(htmlFile);
            var doc = new HtmlDocument();
            doc.Load(htmlRead);

            CollectExternalScripts(doc);
            CollectExternalLinks(doc);

            CreateMht();
            htmlRead.Dispose();
        }

        private void CreateMht()
        {
            var targetFile = _targetFile?? GetMhtFile();
            using (StreamReader rd = new StreamReader(_sourceHtmlFile))
            {
                var item = string.Empty;
                using (StreamWriter wt = new StreamWriter(targetFile))
                {
                    wt.WriteLine("MIME-Version: 1.0");
                    wt.WriteLine("Content-Type: Multipart/related; boundary=\"boundary\"; type=Text/HTML");
                    wt.WriteLine();
                    wt.WriteLine("--boundary");
                    wt.WriteLine("Content-Type: text/html;");
                    wt.WriteLine($"Content-Location: {Path.GetFileName(targetFile)}");
                    wt.WriteLine(rd.ReadToEnd());
                    foreach (var script in _scriptSources.Keys)
                    {
                        if (!string.IsNullOrEmpty(_scriptSources[script]))
                        {
                            //Write external scripts to target.
                            using (StreamReader scriptRd = new StreamReader(_scriptSources[script]))
                            {
                                wt.WriteLine("");
                                wt.WriteLine("--boundary");
                                wt.WriteLine("Content-Type: application/javascript;");
                                wt.WriteLine($"Content-Location: {Path.GetFileName(_scriptSources[script])}");
                                wt.WriteLine("");
                                wt.WriteLine(scriptRd.ReadToEnd());
                            }
                        }
                    }
                    foreach (var externalHtml in _externalResource.Keys)
                    {
                        if (!string.IsNullOrEmpty(_externalResource[externalHtml]))
                        {
                            //Write external links to target.
                            using (StreamReader htmReader = new StreamReader(_externalResource[externalHtml]))
                            {
                                wt.WriteLine("");
                                wt.WriteLine("--boundary");
                                wt.WriteLine("Content-Type: text/html;");
                                wt.WriteLine($"Content-Location: {Path.GetFileName(_externalResource[externalHtml])}");
                                wt.WriteLine("");
                                wt.WriteLine(htmReader.ReadToEnd());
                            }
                        }
                    }
                    wt.WriteLine("");
                    wt.WriteLine("--boundary--");
                }
            }
            OutputFile = targetFile;
        }

        private string GetMhtFile()=>
            Path.Combine(_sourceHtmlDirectory, $"{Path.GetFileNameWithoutExtension(_sourceHtmlFile)}.mht");

        private void CollectExternalLinks(HtmlDocument doc)
        {
            //Collect internal and external links
            var aNodes = doc.DocumentNode.SelectNodes("//a");
            foreach (var aNode in aNodes)
            {
                var hrefAttributeValue = aNode.GetAttributeValue("href", "");
                if (!string.IsNullOrEmpty(hrefAttributeValue))
                {
                    // if (hrefAttributeValue.Contains("#") && !hrefAttributeValue.StartsWith("3D"))
                    // {
                    //     aNode.SetAttributeValue("href", hrefAttributeValue.Replace("#", $"{Path.GetFileName(targetFile)}#"));
                    // }
                    if (hrefAttributeValue.Contains(".htm") || hrefAttributeValue.Contains(".html"))
                    {
                        if (!_externalResource.ContainsKey(hrefAttributeValue))
                            _externalResource.Add(hrefAttributeValue, Directory.GetFiles(_sourceHtmlDirectory, hrefAttributeValue).FirstOrDefault());
                    }
                }
            }
        }

        private void CollectExternalScripts(HtmlDocument doc)
        {
            //Collect Scripts
            var scriptNodes = doc.DocumentNode.SelectNodes("//SCRIPT");
            foreach (var scriptNode in scriptNodes)
            {
                var srcAttributeValue = scriptNode.GetAttributeValue("src", "");
                if (!string.IsNullOrEmpty(srcAttributeValue) && !_scriptSources.ContainsKey(srcAttributeValue))
                    _scriptSources.Add(srcAttributeValue, Directory.GetFiles(_sourceHtmlDirectory, srcAttributeValue).FirstOrDefault());
            }
        }

        private string GetTempHtml() => _targetFile != null ?
            Path.Combine(Path.GetDirectoryName(_targetFile), $"{Path.GetFileNameWithoutExtension(_targetFile)}_temp{Path.GetExtension(_targetFile)}") :
            Path.Combine(Path.GetDirectoryName(_sourceHtmlFile), $"{Path.GetFileNameWithoutExtension(_sourceHtmlFile)}_temp{Path.GetExtension(_sourceHtmlFile)}");
    }
}
