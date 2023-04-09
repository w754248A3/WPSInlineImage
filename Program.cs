using System.Text;
using System.Xml;
using System.IO.Packaging;
using System.Text.RegularExpressions;

internal class Program
{

    static Stream GetXMLStream(Package package, Uri uri)
    {

        var part = package.GetPart(uri);

        var stream = part.GetStream(FileMode.Open, FileAccess.Read);


        var memory = new MemoryStream();

        stream.CopyTo(memory);

        memory.Position = 0;


        stream.Close();

        return memory;

    }

    static T Get元素数据<T>(Stream stream, Func<XmlDocument, T> func)
    {
        var xml = new XmlDocument();

        xml.Load(stream);

        return func(xml);


    }

    static Uri s_workbook = new Uri("/xl/workbook.xml", UriKind.Relative);

    static Uri s_cellimages = new Uri("/xl/cellimages.xml", UriKind.Relative);
    static Uri s_rels_cellimages = new Uri("/xl/_rels/cellimages.xml.rels", UriKind.Relative);
    static Uri s_rels_workbook = new Uri("/xl/_rels/workbook.xml.rels", UriKind.Relative);

    static XmlDocument CreateXml()
    {
        var xml = new XmlDocument();

        var header = xml.CreateXmlDeclaration("1.0", "UTF-8", "yes");

        xml.AppendChild(header);

        return xml;

    }

    static XmlElement F创建元素使用命名空间(XmlDocument xml, string v前缀, string v名字)
    {

        return xml.CreateElement(v前缀, v名字, F根据前缀获取命名空间(v前缀));

    }

    static string F根据前缀获取命名空间(string s)
    {

        if (s == "xdr")
        {
            return "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        }
        else if (s == "r")
        {
            return "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        }
        else if (s == "a")
        {
            return "http://schemas.openxmlformats.org/drawingml/2006/main";
        }
        else if (s == "etc")
        {
            return "http://www.wps.cn/officeDocument/2017/etCustomData";
        }
        else
        {
            throw new ArgumentException("未知命名空间");
        }
    }

    static byte[] F创建新的长ID与短ID映射关系的XML(Action<XmlDocument, XmlElement> action)
    {
        var xml = CreateXml();

        var body = F创建元素使用命名空间(xml, "etc", "cellImages");

        body.SetAttribute("xmlns:xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
        body.SetAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        body.SetAttribute("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        body.SetAttribute("xmlns:etc", "http://www.wps.cn/officeDocument/2017/etCustomData");




        xml.AppendChild(body);

        action(xml, body);

        return F获取二进制XML(xml);
    }

    static void F重新生成长ID与短ID映射关系文件(Package package, List<Cla图片短ID与长ID关系> items)
    {




        var bytes = F创建新的长ID与短ID映射关系的XML((xml, body) =>
        {




            foreach (var item in items)
            {
                XmlElement Create_xdr_nvPicPr()
                {
                    var v_xdr_cNvPr = F创建元素使用命名空间(xml, "xdr", "cNvPr");

                    v_xdr_cNvPr.SetAttribute("id", item.id.ToString());

                    v_xdr_cNvPr.SetAttribute("name", item.v长ID);

                    v_xdr_cNvPr.SetAttribute("descr", item.v描述);


                    var v_xdr_nvPicPr = F创建元素使用命名空间(xml, "xdr", "nvPicPr");

                    v_xdr_nvPicPr.AppendChild(v_xdr_cNvPr);

                    v_xdr_nvPicPr.AppendChild(F创建元素使用命名空间(xml, "xdr", "cNvPicPr"));

                    return v_xdr_nvPicPr;

                }

                XmlElement Create_xdr_blipFill()
                {


                    var v_a_blip = F创建元素使用命名空间(xml, "a", "blip");

                    var v特性 = xml.CreateAttribute("r", "embed", F根据前缀获取命名空间("r"));
                    v特性.Value = $"rId" + item.v短ID;

                    v_a_blip.Attributes.Append(v特性);

                    var v_a_stretch = F创建元素使用命名空间(xml, "a", "stretch");

                    v_a_stretch.AppendChild(F创建元素使用命名空间(xml, "a", "fillRect"));


                    var v_xdr_blipFill = F创建元素使用命名空间(xml, "xdr", "blipFill");

                    v_xdr_blipFill.AppendChild(v_a_blip);

                    v_xdr_blipFill.AppendChild(v_a_stretch);

                    return v_xdr_blipFill;
                }


                XmlElement Create_xdr_spPr()
                {


                    var v_a_off = F创建元素使用命名空间(xml, "a", "off");

                    v_a_off.SetAttribute("x", "0");
                    v_a_off.SetAttribute("y", "0");

                    var v_a_ext = F创建元素使用命名空间(xml, "a", "ext");

                    v_a_ext.SetAttribute("cx", "5657850");
                    v_a_ext.SetAttribute("cy", "10058400");

                    var v_a_xfrm = F创建元素使用命名空间(xml, "a", "xfrm");

                    v_a_xfrm.AppendChild(v_a_off);

                    v_a_xfrm.AppendChild(v_a_ext);

                    var v_a_prstGeom = F创建元素使用命名空间(xml, "a", "prstGeom");

                    v_a_prstGeom.SetAttribute("prst", "rect");

                    v_a_prstGeom.AppendChild(F创建元素使用命名空间(xml, "a", "avLst"));




                    var v_xdr_spPr = F创建元素使用命名空间(xml, "xdr", "spPr");
                    v_xdr_spPr.AppendChild(v_a_xfrm);

                    v_xdr_spPr.AppendChild(v_a_prstGeom);

                    return v_xdr_spPr;

                }

                var v_xdr_pic = F创建元素使用命名空间(xml, "xdr", "pic");


                v_xdr_pic.AppendChild(Create_xdr_nvPicPr());

                v_xdr_pic.AppendChild(Create_xdr_blipFill());

                v_xdr_pic.AppendChild(Create_xdr_spPr());

                var v_etc_cellImage = F创建元素使用命名空间(xml, "etc", "cellImage");

                v_etc_cellImage.AppendChild(v_xdr_pic);


                body.AppendChild(v_etc_cellImage);


            }
        });

        F将数据放入已存在包(package, s_cellimages, "application/vnd.wps-officedocument.cellimage+xml", bytes);


    }

    static byte[] F获取二进制XML(XmlDocument xml)
    {
        byte[] v新的资源文件资源;
        using (var v内存流 = new MemoryStream())
        using (var v编码流 = new StreamWriter(v内存流, Encoding.UTF8))
        {
            xml.Save(v编码流);

            v编码流.Flush();

            v新的资源文件资源 = v内存流.ToArray();

        }

        return v新的资源文件资源;
    }

    static byte[] Create_Cellimages_Xml_Rels(Action<XmlDocument, XmlElement> action)
    {
        var xml = CreateXml();

        var body = xml.CreateElement("Relationships");

        body.SetAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

        xml.AppendChild(body);

        action(xml, body);

        return F获取二进制XML(xml);

    }

    static string F获取图片文件类型(string v扩展名)
    {
        if (v扩展名.Equals(".png", StringComparison.OrdinalIgnoreCase))
        {
            return "image/png";
        }
        else if (v扩展名.Equals(".jpeg", StringComparison.OrdinalIgnoreCase))
        {
            return "image/jpeg";
        }
        else if (v扩展名.Equals(".jpg", StringComparison.OrdinalIgnoreCase))
        {
            return "image/jpeg";
        }
        else
        {
            throw new ArgumentException("改扩展名未知图片类型");
        }
    }


    public record Cla图片信息
    {
        public required int id { get; init; }
        public required string v原始文件名 { get; init; }
        public required string v类型 { get; init; }
        public required string v保存路径 { get; init; }
        public required string v引用路径 { get; init; }
        public required byte[] v图片数据 { get; init; }

        public required string v长ID { get; init; }

    };

    static List<Cla图片信息> F根据图片路径与起始ID生成图片信息(List<Cla图片资源文件数据> v已存在的资源文件, IEnumerable<string> v图片本机路径)
    {
        int v起始id;

        if (v已存在的资源文件.Count == 0)
        {
            v起始id = 0;
        }
        else
        {
            v起始id = v已存在的资源文件.Max(p => p.id);
        }



        var v绝对父路径 = "/xl/media/";
        var v相对父路径 = "media/";

        var vs = new List<Cla图片信息>();

        foreach (var (path, id) in v图片本机路径.Select((path, index) => (path, (v起始id + index) + 1)))
        {
            var v原始文件名 = Path.GetFileNameWithoutExtension(path);

            var v扩展名 = Path.GetExtension(path);

            var v新的文件名字 = $"image{id}{v扩展名}";

            var v保存路径 = Path.Combine(v绝对父路径, v新的文件名字);


            var v引用路径 = Path.Combine(v相对父路径, v新的文件名字);

            var v文件类型 = F获取图片文件类型(v扩展名);
            var v原始图片数据 = File.ReadAllBytes(path);
            var v长ID = "ID_" + Guid.NewGuid().ToString().Replace("-", "");

            vs.Add(new Cla图片信息
            {

                id = id,
                v保存路径 = v保存路径,
                v引用路径 = v引用路径,
                v原始文件名 = v原始文件名,
                v图片数据 = v原始图片数据,
                v类型 = v文件类型,
                v长ID = v长ID



            });

        }

        return vs;

    }

    static void F将数据放入新建包(Package package, Uri uri, string type, byte[] bytes)
    {
        var pack = package.CreatePart(uri, type);

        using (var stream = pack.GetStream(FileMode.Open, FileAccess.Write))
        {
            stream.Write(bytes);
        }
    }

    static void F将数据放入已存在包(Package package, Uri uri, string type, byte[] bytes)
    {

        PackagePart part;
        if (package.PartExists(uri))
        {
            part = package.GetPart(uri);
        }
        else
        {
            part = package.CreatePart(uri, type);
        }


        using (var stream = part.GetStream(FileMode.Open, FileAccess.Write))
        {
            stream.Write(bytes);
        }
    }

    static void F将图片放入图片文件夹(Package package, IEnumerable<Cla图片信息> items)
    {


        foreach (var item in items)
        {

            F将数据放入新建包(package, new Uri(item.v保存路径, UriKind.Relative), item.v类型, item.v图片数据);
        }

    }

    record Cla图片资源文件数据
    {
        public required int id { get; init; }
        public required string target { get; init; }
    }

    static List<Cla图片资源文件数据> F获取已存在的图片资源文件信息(Package package)
    {
        var uri = s_rels_cellimages;

        if (!package.PartExists(uri))
        {
            return new List<Cla图片资源文件数据>();
        }


        using (var stream = GetXMLStream(package, uri))
        {
            return Get元素数据(stream, (xml) =>
                    {


                        return xml.GetElementsByTagName("Relationship").OfType<XmlElement>()
                        .Select(p => new Cla图片资源文件数据
                        {
                            id = F解析出短ID(p.GetAttribute("Id")),
                            target = p.GetAttribute("Target")
                        })
                        .ToList();


                    });
        }


    }

    static void F重新生成图片资源文件(Package package, List<Cla图片资源文件数据> v已存在图片信息, List<Cla图片信息> v欲添加图片信息)
    {



        var items = v已存在图片信息.Concat(v欲添加图片信息.Select(p => new Cla图片资源文件数据 { id = p.id, target = p.v引用路径 }));

        var bytes = Create_Cellimages_Xml_Rels((xml, e) =>
            {
                foreach (var item in items)
                {
                    var new_e = xml.CreateElement("Relationship");
                    new_e.SetAttribute("Id", $"rId{item.id}");
                    new_e.SetAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

                    new_e.SetAttribute("Target", item.target);

                    e.AppendChild(new_e);
                }
            });




        F将数据放入已存在包(package, s_rels_cellimages, "application/vnd.openxmlformats-package.relationships+xml", bytes);




    }

    public record Cla图片短ID与长ID关系
    {
        public required int id { get; init; }
        public required int v短ID { get; init; }
        public required string v长ID { get; init; }

        public required string v描述 { get; init; }
    }

    static int F解析出短ID(string s)
    {
        return int.Parse(Regex.Match(s, @"\d+").Value);
    }

    static List<Cla图片短ID与长ID关系> F获取已存在的图片短ID与长ID关系(Package package)
    {
        var uri = s_cellimages;
        if (!package.PartExists(uri))
        {
            return new List<Cla图片短ID与长ID关系>();
        }


        using (var stream = GetXMLStream(package, uri))
        {
            var res = Get元素数据(stream, (xml) =>
               {
                   var vs = xml.GetElementsByTagName("etc:cellImage")
                   .OfType<XmlElement>();


                   var res = new List<Cla图片短ID与长ID关系>();
                   foreach (var item in vs)
                   {
                       var v = item.GetElementsByTagName("xdr:cNvPr").OfType<XmlElement>().First();

                       var id = int.Parse(v.GetAttribute("id"));

                       var v长ID = v.GetAttribute("name");

                       var v描述 = v.GetAttribute("descr");

                       v = item.GetElementsByTagName("a:blip").OfType<XmlElement>().First();

                       var v短ID = F解析出短ID(v.GetAttribute("r:embed"));


                       res.Add(new Cla图片短ID与长ID关系
                       {
                           id = id,
                           v长ID = v长ID,
                           v短ID = v短ID,
                           v描述 = v描述

                       });

                   }

                   return res;
               });
            return res;
        }




    }


    static List<Cla图片短ID与长ID关系> F合并图片长ID与短ID映射关系(List<Cla图片短ID与长ID关系> v已存在的映射关系, List<Cla图片信息> v新添加的映射关系)
    {


        int v最大映射文件id;
        if (v已存在的映射关系.Count == 0)
        {
            v最大映射文件id = 0;
        }
        else
        {
            v最大映射文件id = v已存在的映射关系.Max(p => p.id);
        }


        return v已存在的映射关系.Concat(v新添加的映射关系.Select((p, index) => new Cla图片短ID与长ID关系
        {

            id = v最大映射文件id + index + 1,

            v描述 = p.v原始文件名,
            v短ID = p.id,
            v长ID = p.v长ID
        })).ToList();
    }

    static void F将映射文件注册(Package package)
    {
        var uri = new Uri("cellimages.xml", UriKind.Relative);
        var part = package.GetPart(s_workbook);

        var v = part.GetRelationships()
        .Any(p => p.TargetUri == uri);
        Console.WriteLine(v);
        if (!v)
        {
            part.CreateRelationship(uri, TargetMode.Internal, "http://www.wps.cn/officeDocument/2020/cellImage");
        }



        foreach (var item in part.GetRelationships())
        {
            Console.WriteLine($"{item.SourceUri} {item.TargetMode} {item.RelationshipType} {item.TargetUri}");
        }

    }
    private static void Main(string[] args)
    {
        var p1 = @"C:\Users\PC\Desktop\教程\test2.xlsx";

        var p2 = @"C:\Users\PC\Desktop\教程\test.xlsx";

        File.Copy(p1, p2, true);

        var package = System.IO.Packaging.Package.Open(p2);

        var v图片文件夹 = @"C:\Users\PC\Desktop\教程\图片";

        var v已存在图片信息 = F获取已存在的图片资源文件信息(package);
        foreach (var item in v已存在图片信息)
        {
            Console.WriteLine(item);
        }



        var v添加的图片信息 = F根据图片路径与起始ID生成图片信息(v已存在图片信息, Directory.GetFiles(v图片文件夹));

        foreach (var item in v添加的图片信息)
        {
            Console.WriteLine(item);
        }

        F将图片放入图片文件夹(package, v添加的图片信息);

        F重新生成图片资源文件(package, v已存在图片信息, v添加的图片信息);

        var v已存在的图片映射 = F获取已存在的图片短ID与长ID关系(package);


        F重新生成长ID与短ID映射关系文件(package, F合并图片长ID与短ID映射关系(v已存在的图片映射, v添加的图片信息));

        F将映射文件注册(package);
        package.Close();

    }

}