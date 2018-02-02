 using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;

namespace TableMaster
{
    public class XMLOperator
    {
        public static List<String> ReadXML(String filename,String c_name)
        {
            XDocument document = XDocument.Load(filename);
            XElement root = document.Root;
            XElement ele = root.Element(c_name);
            List<String> list = new List<String>();
            foreach (XElement item in ele.Elements())
            {
                list.Add(item.Value);
            }
            return list;
        }

        public static void ReadXML(String filename, ListBox tar)
        {
            XDocument document = XDocument.Load(filename);
            XElement root = document.Root;
            XElement ele = root.Element(tar.Name);
            List<String> list = new List<String>();
            foreach (XElement item in ele.Elements())
            {
                ListBoxItem n_item = new ListBoxItem();
                n_item.Content = item.Value;
                tar.Items.Add(n_item);
            }
        }

        public static void InitializeContainer(String filename,Window window)
        {
            if (!File.Exists(filename))
            {
                return;
            }
            XDocument document = XDocument.Load(filename);
            XElement root = document.Root;
            foreach (XElement ele in root.Elements())
            {
                object tar = window.FindName(ele.Name.LocalName);
                if(ele.Elements().ElementAt(0).Name.LocalName == "Content")
                {
                    ((TextBox)tar).Text = ele.Value;
                }
                else if(ele.Elements().ElementAt(0).Name.LocalName == "Content0")
                {
                    foreach (XElement item in ele.Elements())
                    {
                        ListBoxItem n_item = new ListBoxItem();
                        n_item.Content = item.Value;
                        ((ListBox)tar).Items.Add(n_item);
                    }
                }
            }
        }

        public static void AddIntoXML(String fileName,String c_name,String content)
        {
            XDocument document = null;
            XElement root = null;
            if (File.Exists(fileName))
            {
                document = XDocument.Load(fileName);
                root = document.Root;
            }
            else
            {
                document = new XDocument();
                root = new XElement("Components");
                document.Add(root);
            }
            if (root.Element(c_name) == null)
            {
                XElement n_ele = new XElement(c_name);
                n_ele.SetElementValue("Content", content);
                root.Add(n_ele);
            }
            else
            {
                XElement n_ele = root.Element(c_name);
                n_ele.SetElementValue("Content", content);
            }
            document.Save(fileName);
        }

        public static void AddIntoXML(String fileName, ListBox c)
        {
            XDocument document = null;
            XElement root = null;
            if (File.Exists(fileName))
            {
                document = XDocument.Load(fileName);
                root = document.Root;
            }
            else
            {
                document = new XDocument();
                root = new XElement("Components");
                document.Add(root);
            }

            XElement n_ele = null;
            if (root.Element(c.Name) == null)
            {
                n_ele = new XElement(c.Name);
            }
            else
            {
                n_ele = root.Element(c.Name);
                n_ele.RemoveAll();
            }

            for(int i = 0; i<c.Items.Count; i++)
            {
                ListBoxItem item = (ListBoxItem)c.Items[i];
                n_ele.SetElementValue("Content" + i, item.Content);
            }

            root.Add(n_ele);
            document.Save(fileName);
        }
    }
}
