using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using WordPadClone.Formats;

internal static class P
{
    private static void Main()
    {
        string root = "tmp-regression";
        Directory.CreateDirectory(root);

        string traversal = Path.Combine(root, "zip-traversal.zip");
        ZipArchiveLite.Write(traversal, new List<ZipEntryData>
        {
            new ZipEntryData("../evil.txt", "x", true)
        });
        ExpectInvalid(delegate { ZipArchiveLite.Read(traversal); }, "zip traversal");

        string duplicate = Path.Combine(root, "zip-duplicate.zip");
        ZipArchiveLite.Write(duplicate, new List<ZipEntryData>
        {
            new ZipEntryData("word/document.xml", "a", true),
            new ZipEntryData("word/document.xml", "b", true)
        });
        ExpectInvalid(delegate { ZipArchiveLite.Read(duplicate); }, "zip duplicate");

        string utf8 = Path.Combine(root, "bom-utf8.txt");
        File.WriteAllBytes(utf8, new byte[] { 0xEF, 0xBB, 0xBF, (byte)'h', (byte)'i' });
        if (TextFile.Load(utf8) != "hi") throw new Exception("UTF-8 BOM");

        string utf16 = Path.Combine(root, "bom-utf16.txt");
        File.WriteAllBytes(utf16, new byte[] { 0xFF, 0xFE, (byte)'h', 0, (byte)'i', 0 });
        if (TextFile.Load(utf16) != "hi") throw new Exception("UTF-16 BOM");

        StringBuilder deep = new StringBuilder();
        for (int i = 0; i < 300; i++) deep.Append('{');
        deep.Append("x");
        for (int i = 0; i < 300; i++) deep.Append('}');
        ExpectInvalid(delegate { RtfReader.Read("{\\rtf1" + deep + "}"); }, "RTF nesting");

        string clean = Path.Combine(root, "security-clean.docx");
        if (File.Exists(clean)) File.Delete(clean);
        if (File.Exists(clean + ".bak")) File.Delete(clean + ".bak");
        DocumentModel model = new DocumentModel();
        Paragraph p = new Paragraph();
        p.Inlines.Add(new TextInline("security", new RunFormat("Calibri", 22)));
        model.Paragraphs.Add(p);
        DocumentFormats.SaveDocx(clean, model);
        Dictionary<string, byte[]> package = ZipArchiveLite.Read(clean);
        package["word/vbaProject.bin"] = new byte[] { 1, 2, 3 };
        package["word/activeX/control1.xml"] = Encoding.UTF8.GetBytes("<control/>");
        string risky = Path.Combine(root, "security-macro.docx");
        List<ZipEntryData> entries = new List<ZipEntryData>();
        foreach (KeyValuePair<string, byte[]> entry in package)
            entries.Add(new ZipEntryData(entry.Key, entry.Value, true));
        ZipArchiveLite.Write(risky, entries);
        DocumentModel reopened = DocxFormat.Read(risky);
        bool macroWarning = false;
        bool activeXWarning = false;
        for (int i = 0; i < reopened.SecurityWarnings.Count; i++)
        {
            if (reopened.SecurityWarnings[i].IndexOf("macro", StringComparison.OrdinalIgnoreCase) >= 0)
                macroWarning = true;
            if (reopened.SecurityWarnings[i].IndexOf("ActiveX", StringComparison.OrdinalIgnoreCase) >= 0)
                activeXWarning = true;
        }
        if (!macroWarning || !activeXWarning) throw new Exception("security warning missing");
        foreach (string key in reopened.PreservedParts.Keys)
            if (key.StartsWith("word/activeX/", StringComparison.OrdinalIgnoreCase)
                || key.StartsWith("word/vbaProject.bin", StringComparison.OrdinalIgnoreCase))
                throw new Exception("active content preserved");

        string odt = Path.Combine(root, "security-dtd.odt");
        ZipArchiveLite.Write(odt, new List<ZipEntryData>
        {
            new ZipEntryData("content.xml",
                "<?xml version=\"1.0\"?><!DOCTYPE x [<!ENTITY xxe SYSTEM \"file:///c:/windows/win.ini\">]><x>&xxe;</x>",
                true)
        });
        ExpectAny(delegate { OdtFormat.Read(odt); }, "ODT DTD");

        byte[] hugePng = new byte[24];
        hugePng[0] = 0x89; hugePng[1] = 0x50; hugePng[2] = 0x4E; hugePng[3] = 0x47;
        hugePng[16] = 0; hugePng[17] = 1; hugePng[18] = 0x86; hugePng[19] = 0xA0;
        hugePng[20] = 0; hugePng[21] = 1; hugePng[22] = 0x86; hugePng[23] = 0xA0;
        ExpectInvalid(delegate { ImageUtil.LoadImage(hugePng, "png", 100, 100); }, "image dimensions");

        byte[] negativePng = (byte[])hugePng.Clone();
        negativePng[16] = 0xFF; negativePng[17] = 0xFF; negativePng[18] = 0xFF; negativePng[19] = 0xFF;
        ExpectInvalid(delegate { ImageUtil.LoadImage(negativePng, "png", 100, 100); }, "negative PNG dimensions");

        string externalRelPath = Path.Combine(root, "external-rel.docx");
        if (File.Exists(externalRelPath)) File.Delete(externalRelPath);
        if (File.Exists(externalRelPath + ".bak")) File.Delete(externalRelPath + ".bak");
        DocumentFormats.SaveDocx(externalRelPath, model);
        Dictionary<string, byte[]> externalPackage = ZipArchiveLite.Read(externalRelPath);
        string rels = Encoding.UTF8.GetString(externalPackage["word/_rels/document.xml.rels"]);
        rels = rels.Replace("</Relationships>",
            "<Relationship Id=\"rIdExternal\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"file:///c:/secret.png\" TargetMode=\"External\" /></Relationships>");
        externalPackage["word/_rels/document.xml.rels"] = Encoding.UTF8.GetBytes(rels);
        List<ZipEntryData> externalEntries = new List<ZipEntryData>();
        foreach (KeyValuePair<string, byte[]> entry in externalPackage)
            externalEntries.Add(new ZipEntryData(entry.Key, entry.Value, true));
        ZipArchiveLite.Write(externalRelPath, externalEntries);
        DocumentModel externalModel = DocxFormat.Read(externalRelPath);
        for (int i = 0; i < externalModel.PreservedRelationships.Count; i++)
            if (externalModel.PreservedRelationships[i].Target.IndexOf("secret.png", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new Exception("external relationship was preserved");
        TestNestedDocx(root);
        TestNestedOdt(root);
        Console.WriteLine("SECURITY_SMOKE=PASS");
    }

    private static void TestNestedDocx(string root)
    {
        DocumentModel model = new DocumentModel();
        Paragraph p = new Paragraph();
        p.Inlines.Add(new TextInline("x", new RunFormat("Calibri", 22)));
        model.Paragraphs.Add(p);
        string path = Path.Combine(root, "deep-docx.docx");
        if (File.Exists(path)) File.Delete(path);
        if (File.Exists(path + ".bak")) File.Delete(path + ".bak");
        DocumentFormats.SaveDocx(path, model);
        Dictionary<string, byte[]> package = ZipArchiveLite.Read(path);
        StringBuilder body = new StringBuilder();
        for (int i = 0; i < 200; i++) body.Append("<w:sdt><w:sdtContent>");
        body.Append("<w:p><w:r><w:t>x</w:t></w:r></w:p>");
        for (int i = 0; i < 200; i++) body.Append("</w:sdtContent></w:sdt>");
        string xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>"
            + body + "</w:body></w:document>";
        package["word/document.xml"] = Encoding.UTF8.GetBytes(xml);
        List<ZipEntryData> entries = new List<ZipEntryData>();
        foreach (KeyValuePair<string, byte[]> entry in package)
            entries.Add(new ZipEntryData(entry.Key, entry.Value, true));
        ZipArchiveLite.Write(path, entries);
        ExpectInvalid(delegate { DocxFormat.Read(path); }, "DOCX depth");
    }

    private static void TestNestedOdt(string root)
    {
        DocumentModel model = new DocumentModel();
        Paragraph p = new Paragraph();
        p.Inlines.Add(new TextInline("x", new RunFormat("Calibri", 22)));
        model.Paragraphs.Add(p);
        string path = Path.Combine(root, "deep-odt.odt");
        if (File.Exists(path)) File.Delete(path);
        if (File.Exists(path + ".bak")) File.Delete(path + ".bak");
        DocumentFormats.SaveOdt(path, model);
        Dictionary<string, byte[]> package = ZipArchiveLite.Read(path);
        StringBuilder body = new StringBuilder();
        for (int i = 0; i < 200; i++) body.Append("<text:section>");
        body.Append("<text:p>x</text:p>");
        for (int i = 0; i < 200; i++) body.Append("</text:section>");
        string xml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"><office:body><office:text>"
            + body + "</office:text></office:body></office:document-content>";
        package["content.xml"] = Encoding.UTF8.GetBytes(xml);
        List<ZipEntryData> entries = new List<ZipEntryData>();
        foreach (KeyValuePair<string, byte[]> entry in package)
            entries.Add(new ZipEntryData(entry.Key, entry.Value, true));
        ZipArchiveLite.Write(path, entries);
        ExpectInvalid(delegate { OdtFormat.Read(path); }, "ODT depth");
    }
    private static void ExpectInvalid(Action action, string name)
    {
        try
        {
            action();
        }
        catch (InvalidDataException)
        {
            return;
        }
        throw new Exception(name + " was accepted");
    }

    private static void ExpectAny(Action action, string name)
    {
        try
        {
            action();
        }
        catch (Exception)
        {
            return;
        }
        throw new Exception(name + " was accepted");
    }
}
