using System.IO;
using System.Windows.Forms;

namespace DragDropTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            //display formats available
            label1.Text = "Formats:\n";
            foreach (var format in e.Data.GetFormats())
            {
                label1.Text += "    " + format + "\n";
            }

            //ensure FileGroupDescriptor is present before allowing drop
            if (e.Data.GetDataPresent("FileGroupDescriptor"))
            {
                e.Effect = DragDropEffects.All;
            }
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            //wrap standard IDataObject in OutlookDataObject
            var dataObject = new OutlookDataObject(e.Data);

            //get the names and data streams of the files dropped
            var filenames = (string[])dataObject.GetData("FileGroupDescriptorW");
            var filestreams = (MemoryStream[])dataObject.GetData("FileContents");

            label1.Text += "Files:\n";
            for (var fileIndex = 0; fileIndex < filenames.Length; fileIndex++)
            {
                //use the fileindex to get the name and data stream
                var filename = filenames[fileIndex];
                var filestream = filestreams[fileIndex];
                label1.Text += "    " + filename + "\n";

                //save the file stream using its name to the application path
                var outputStream = File.Create(filename);
                filestream.WriteTo(outputStream);
                outputStream.Close();
            }
        }
    }
}