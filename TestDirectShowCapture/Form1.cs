using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using DirectShowCapture;
using DirectShowLib;

namespace TestDirectShowCapture
{
    public partial class Form1 : Form
    {
        Dictionary<string, DsDevice> dicDevices = new Dictionary<string, DsDevice>();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DsDevice[] devices = DSCapture.GetCaptureDevices();

            dicDevices.Clear();

            foreach (DsDevice device in devices)
            {
                Console.WriteLine(device.Name);
                dicDevices.Add(device.Name, device);
            }
        }

       DSCapture capture = null;
        string deviceName = "AVerMedia GC551 Video Capture";
        bool isRunCapture = false;

        private void button2_Click(object sender, EventArgs e)
        {
            if (isRunCapture == true) return;

            if (capture == null)
            {
                capture = new DSCapture(dicDevices[deviceName], panel1.Handle, panel1.ClientSize);
            }

            capture.Play(true);

            isRunCapture = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (capture == null) return;
            if (isRunCapture == false) return;

            capture.Stop();

            capture.Dispose();
            capture = null;
            Task.Run(() => GC.Collect());

            isRunCapture = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Bitmap bitmap = capture.Capture();
            //Bitmap bitmap = capture.Capture(new Size(640, 360));
            //Bitmap bitmap = capture.Capture(new Rectangle(0, 0, 640, 360));

            Clipboard.SetImage(bitmap);

            // MTA環境でクリップボードにコピー
            //Bitmap bitmap = capture.Capture();
            //Thread thread = new Thread(() => Clipboard.SetImage(bitmap));
            //thread.SetApartmentState(ApartmentState.STA);
            //thread.Start();
            //thread.Join();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (capture.ByteCount != 3) throw new Exception("非対応");

            byte[] buffer = capture.GetBuffer();

            Bitmap bitmap = toBitmap(buffer);

            Clipboard.SetImage(bitmap);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (capture.ByteCount != 3) throw new Exception("非対応");

            byte[] buffer = capture.GetBufferEx();

            Bitmap bitmap = toBitmap(buffer);

            Clipboard.SetImage(bitmap);
        }

        private Bitmap toBitmap(byte[] buffer)
        {
            Bitmap bitmap = new Bitmap(capture.Width, capture.Height, PixelFormat.Format24bppRgb);
            BitmapData bitmapData = bitmap.LockBits(
            new Rectangle(0, 0, bitmap.Width, bitmap.Height),
            ImageLockMode.WriteOnly,
            PixelFormat.Format24bppRgb);
            Marshal.Copy(buffer, 0, bitmapData.Scan0, buffer.Length);
            bitmap.UnlockBits(bitmapData);

            return bitmap;
        }
    }
}
