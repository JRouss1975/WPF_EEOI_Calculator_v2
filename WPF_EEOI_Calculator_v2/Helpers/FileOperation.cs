using System;
using System.Xml.Serialization;
using System.IO;
using System.Windows;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;
using System.Xml;

namespace WPF_EEOI_Calculator_v2
{
    public static class FileOperation
    {
        public static string FileName = "";

        public static T OpenBinaryObject<T>() where T : class
        {
            T obj = null;

            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".dat";
            dlg.Filter = "dat files (*.dat)|*.dat|All files (*.*)|*.*";

            // Display OpenFileDialog by calling ShowDialog method 
            bool? result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                try
                {
                    BinaryFormatter xmlFormat = new BinaryFormatter();
                    using (Stream fStream = new FileStream(dlg.FileName, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        fStream.Position = 0;
                        obj = (T)xmlFormat.Deserialize(fStream);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("File NOT loaded!");
                }
            }
            return obj;
        }

        public static void SaveObjectToBinary<T>(T obj)
        {
            // Create SaveFileDialog 
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".dat";
            dlg.Filter = "dat files (*.dat)|*.dat|All files (*.*)|*.*";

            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();
            try
            {
                BinaryFormatter binFormat = new BinaryFormatter();
                using (Stream fStream = new FileStream(dlg.FileName, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    fStream.Position = 0;
                    binFormat.Serialize(fStream, obj);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        public static void SaveObjectToXML<T>(T obj)
        {

            try
            {
                XmlSerializer ds = new XmlSerializer(typeof(T));
                XmlWriterSettings settings = new XmlWriterSettings() { Indent = true };
                using (XmlWriter fStream = XmlWriter.Create(FileName, settings))
                {
                    ds.Serialize(fStream, obj);
                    MessageBox.Show("File:" + FileOperation.FileName + " saved!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }

        public static void SaveAsObjectToXML<T>(T obj)
        {
            // Create SaveFileDialog 
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xml";
            dlg.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*";

            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();
            try
            {
                XmlSerializer ds = new XmlSerializer(typeof(T));
                XmlWriterSettings settings = new XmlWriterSettings() { Indent = true };
                using (XmlWriter fStream = XmlWriter.Create(dlg.FileName, settings))
                {
                    ds.Serialize(fStream, obj);
                    MessageBox.Show("File:" + dlg.FileName + " saved!");
                    FileOperation.FileName = dlg.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }



        public static T OpenXMLObject<T>() where T : class
        {
            T obj = null;

            // Create SaveFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xml";
            dlg.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*";

            // Display OpenFileDialog by calling ShowDialog method 
            bool? result = dlg.ShowDialog();
            try
            {

                //var ds = new NetDataContractSerializer();
                XmlSerializer ds = new XmlSerializer(typeof(T));
                using (XmlReader fStream = XmlReader.Create(dlg.FileName))
                {
                    obj = (T)ds.Deserialize(fStream);
                    FileName = dlg.FileName;
                }
                //obj = (T)ds.ReadObject(fStream);
            }
            catch (Exception)
            {
                MessageBox.Show("File could not be loaded.");
            }
            return obj;
        }
    }
}