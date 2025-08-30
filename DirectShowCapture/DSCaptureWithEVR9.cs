using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using DirectShowLib;

namespace DirectShowCapture
{
    public class DSCaptureWithEVR9 : IDisposable, ISampleGrabberCB
    {
        /// <summary>
        /// バッファを移動する
        /// </summary>
        /// <param name="Destination">移動先のメモリへのポインタ</param>
        /// <param name="Source">移動元のメモリへのポインタ</param>
        /// <param name="Length">移動するメモリのバイト数</param>
        [DllImport("Kernel32.dll", EntryPoint = "RtlMoveMemory")]
        private static extern void MoveMemory(IntPtr Destination, IntPtr Source, [MarshalAs(UnmanagedType.U4)] int Length);

        const int WMGraphNotify = 0x0400 + 13;

        private Size _renderingSize;
        public Size RenderingSize
        {
            get
            {
                return _renderingSize;
            }
            set
            {
                _renderingSize = value;

                int pWidth, pHeight, hr;
                hr = this.basicVideo.GetVideoSize(out pWidth, out pHeight);
                DsError.ThrowExceptionForHR(hr);
                Console.WriteLine("{{Width:{0}, Height:{1}}}", pWidth, pHeight);

                Rectangle windowPos = new Rectangle();
                if (RenderingSize.Width < RenderingSize.Height)
                {
                    // 縦長
                    windowPos.X = 0;
                    windowPos.Width = value.Width;
                    windowPos.Height = value.Width * pHeight / pWidth;
                    windowPos.Y = (value.Height - (windowPos.Height)) / 2;
                }
                else
                {
                    // 横長
                    windowPos.Y = 0;
                    windowPos.Height = value.Height;
                    windowPos.Width = value.Height * pWidth / pHeight;
                    windowPos.X = (value.Width - (windowPos.Width)) / 2;
                }
                hr = this.videoWindow.SetWindowPosition(
                    windowPos.X,
                    windowPos.Y,
                    windowPos.Width,
                    windowPos.Height);
                DsError.ThrowExceptionForHR(hr);
            }
        }

        IGraphBuilder graphBuilder = null;
        IMediaEventEx mediaEventEx = null;
        IMediaControl mediaControl = null;
        IVideoWindow videoWindow = null;
        IBasicVideo basicVideo = null;

        ISampleGrabber sampleGrabber = null;

        ICaptureGraphBuilder2 captureGraphBuilder;

        IntPtr buffer = IntPtr.Zero;

        private int width = 0;
        public int Width
        {
            get
            {
                return width;
            }
        }
        private int height = 0;
        public int Height
        {
            get
            {
                return height;
            }
        }
        private int byteCount = 0;
        public int ByteCount
        {
            get
            {
                return byteCount;
            }
        }
        private int stride = 0;
        public int Stride
        {
            get
            {
                return stride;
            }
        }

        /// <summary>
        /// 映像キャプチャデバイスの一覧を取得します。
        /// </summary>
        /// <returns></returns>
        public static DsDevice[] GetCaptureDevices()
        {
            return DsDevice.GetDevicesOfCat(FilterCategory.VideoInputDevice);
        }

        /// <summary>
        /// キャプチャフィルタの取得
        /// </summary>
        /// <param name="device">開きたいデバイス</param>
        /// <returns>キャプチャフィルタ</returns>
        private static IBaseFilter GetCaptureDevice(DsDevice device)
        {
            object source;
            Guid filterId = typeof(IBaseFilter).GUID;

            device.Mon.BindToObject(null, null, ref filterId, out source);

            return (IBaseFilter)source;
        }

        public DSCaptureWithEVR9(DsDevice device, IntPtr Handle, Size size) : base()
        {
            //Console.WriteLine("Opening :{0}", device.Name);
            // 初期化
            this.graphBuilder = (IGraphBuilder)new FilterGraph();
            this.captureGraphBuilder = (ICaptureGraphBuilder2)new CaptureGraphBuilder2();

            // キャプチャフィルタの作成
            IBaseFilter videoCaptureFilter = GetCaptureDevice(device);

            // キャプチャグラフの初期化
            int hr = captureGraphBuilder.SetFiltergraph(graphBuilder);
            DsError.ThrowExceptionForHR(hr);

            hr = graphBuilder.AddFilter(videoCaptureFilter, "Video Capture");
            DsError.ThrowExceptionForHR(hr);

            // サンプルグラバの初期化
            this.sampleGrabber = (ISampleGrabber)new SampleGrabber();

            AMMediaType mediaType = new AMMediaType();
            mediaType.majorType = MediaType.Video;        // メジャータイプ
            mediaType.subType = MediaSubType.RGB24;       // サブタイプ
            mediaType.formatType = FormatType.VideoInfo;  // フォーマットタイプ

            // メディア タイプを設定する
            hr = sampleGrabber.SetMediaType(mediaType);
            DsError.ThrowExceptionForHR(hr);

            DsUtils.FreeAMMediaType(mediaType);

            hr = graphBuilder.AddFilter((IBaseFilter)sampleGrabber, "Sample Grabber");
            DsError.ThrowExceptionForHR(hr);

            // ビデオキャプチャフィルタをレンダリングフィルタに接続
            hr = this.captureGraphBuilder.RenderStream(
                PinCategory.Capture,
                MediaType.Video,
                videoCaptureFilter,
                (IBaseFilter)sampleGrabber,
                null);
            DsError.ThrowExceptionForHR(hr);

            // 解放
            System.Runtime.InteropServices.Marshal.ReleaseComObject(videoCaptureFilter);

            // フィルタグラフの設定
            this.mediaEventEx = (IMediaEventEx)graphBuilder;
            this.mediaControl = (IMediaControl)graphBuilder;
            this.videoWindow = this.graphBuilder as IVideoWindow;
            this.basicVideo = this.graphBuilder as IBasicVideo;

            this.mediaEventEx.SetNotifyWindow(Handle, WMGraphNotify, IntPtr.Zero);

            hr = this.videoWindow.put_Owner(Handle);
            DsError.ThrowExceptionForHR(hr);
            hr = this.videoWindow.put_WindowStyle(WindowStyle.Child | WindowStyle.ClipChildren | WindowStyle.ClipSiblings);
            DsError.ThrowExceptionForHR(hr);

            this.RenderingSize = size;
        }

        public void Play(bool captureMode)
        {
            int hr = 0;

            sampleGrabber.SetOneShot(false);
            if (captureMode)
            {
                sampleGrabber.SetBufferSamples(true);
                sampleGrabber.SetCallback(this, 1);
            }
            else
            {
                sampleGrabber.SetBufferSamples(false);
                sampleGrabber.SetCallback(null, 0);
            }

            AMMediaType mediaType = new AMMediaType();

            // サンプル グラバからメディア タイプを取得する
            hr = sampleGrabber.GetConnectedMediaType(mediaType);
            DsError.ThrowExceptionForHR(hr);

            // サイズ情報を取得する
            VideoInfoHeader videoInfoHeader
                = (VideoInfoHeader)Marshal.PtrToStructure(mediaType.formatPtr, typeof(VideoInfoHeader));
            width = videoInfoHeader.BmiHeader.Width;
            height = videoInfoHeader.BmiHeader.Height;
            byteCount = videoInfoHeader.BmiHeader.BitCount / 8;
            stride = width * byteCount;
            DsUtils.FreeAMMediaType(mediaType);

            buffer = Marshal.AllocCoTaskMem(height * stride);

            // スタート
            hr = this.mediaControl.Run();
            DsError.ThrowExceptionForHR(hr);
        }

        public void SetCaptureMode(bool captureMode)
        {
            if (captureMode)
            {
                sampleGrabber.SetBufferSamples(true);
                sampleGrabber.SetCallback(this, 1);
                //Console.WriteLine("capture mode active");
            }
            else
            {
                sampleGrabber.SetBufferSamples(false);
                sampleGrabber.SetCallback(null, 0);
                //Console.WriteLine("capture mode deactive");
            }
        }

        int ISampleGrabberCB.SampleCB(double SampleTime, IMediaSample pSample)
        {
            //Console.WriteLine("SampleCB");
            return 0;
        }

        /// <summary>
        /// サンプリング終了時に呼び出されるコールバック
        /// </summary>
        /// <param name="sampleTime">サンプルの開始時間</param>
        /// <param name="pBuffer">サンプル データを含むバッファへのポインタ</param>
        /// <param name="bufferLength">サンプル データを含むバッファのバイト数</param>
        /// <returns></returns>
        int ISampleGrabberCB.BufferCB(double sampleTime, IntPtr pBuffer, int bufferLength)
        {
            //Console.WriteLine("BufferCB");
            // バッファを保存する
            MoveMemory(this.buffer, pBuffer, bufferLength);

            return 0;
        }

        public void Stop()
        {
            // ストップ
            int hr = this.mediaControl.Stop();
            DsError.ThrowExceptionForHR(hr);

            if (buffer == IntPtr.Zero) return;

            System.Runtime.InteropServices.Marshal.FreeCoTaskMem(buffer);
            buffer = IntPtr.Zero;
        }

        /// <summary>
        /// Bitmapに変換します
        /// </summary>
        /// <returns></returns>
        public Bitmap Capture()
        {
            if (buffer == IntPtr.Zero) return null;

            // ビットマップの作成
            System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(
                width,                      // Bitmapの幅
                height,                     // Bitmapの高さ
                stride,                     // スキャン ラインの間のバイト オフセット数
                PixelFormat.Format24bppRgb, // カラーデータの形式
                buffer                      // ピクセル データを格納しているバッファ
                );

            // 上下を反転する
            bitmap.RotateFlip(System.Drawing.RotateFlipType.RotateNoneFlipY);

            return bitmap;
        }

        /// <summary>
        /// 指定したサイズのBitmapに変換します
        /// </summary>
        /// <returns></returns>
        public Bitmap Capture(Size size)
        {
            if (buffer == IntPtr.Zero) return null;

            Bitmap bitmap = new Bitmap(size.Width, size.Height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.DrawImage(Capture(), 0, 0, size.Width, size.Height);
            }

            return bitmap;
        }

        /// <summary>
        /// 指定した範囲をBitmapに変換します
        /// </summary>
        /// <returns></returns>
        public Bitmap Capture(Rectangle rectangle)
        {
            if (buffer == IntPtr.Zero) return null;

            return Capture().Clone(rectangle, PixelFormat.Format24bppRgb);
        }

        /// <summary>
        /// Bitmap化せずに直接バッファをコピーする<br></br>
        /// データが上下反転しているので注意
        /// </summary>
        /// <returns></returns>
        public byte[] GetBuffer()
        {
            if (buffer == IntPtr.Zero) return null;

            byte[] array = new byte[height * stride];
            Marshal.Copy(buffer, array, 0, array.Length);
            return array;
        }

        /// <summary>
        /// Bitmap化せずに直接バッファをコピーし<br></br>
        /// データの上下反転を正す処理をします<br></br>
        /// ただしEx付きの方が処理は遅いです。
        /// </summary>
        /// <returns></returns>
        public byte[] GetBufferEx()
        {
            if (buffer == IntPtr.Zero) return null;

            byte[] array = new byte[height * stride];
            Marshal.Copy(buffer, array, 0, array.Length);

            byte[] result = new byte[height * stride];
            int length = result.Length;
            int offset = 0;
            while (offset < length)
            {
                Buffer.BlockCopy(array, (length - offset - stride), result, offset, stride);
                offset += stride;
            }
            return result;
        }

        public void Dispose()
        {
            if (sampleGrabber != null) { Marshal.ReleaseComObject(sampleGrabber); sampleGrabber = null; }
            if (basicVideo != null) { Marshal.ReleaseComObject(basicVideo); basicVideo = null; }
            if (videoWindow != null) { Marshal.ReleaseComObject(videoWindow); videoWindow = null; }
            if (mediaControl != null) { Marshal.ReleaseComObject(mediaControl); mediaControl = null; }
            if (mediaEventEx != null) { Marshal.ReleaseComObject(mediaEventEx); mediaEventEx = null; }
            if (graphBuilder != null) { Marshal.ReleaseComObject(graphBuilder); graphBuilder = null; }
        }

        void IDisposable.Dispose()
        {
            if (this.graphBuilder != null) Marshal.ReleaseComObject(this.graphBuilder);
            if (this.captureGraphBuilder != null) Marshal.ReleaseComObject(this.captureGraphBuilder);
            if (this.sampleGrabber != null) Marshal.ReleaseComObject(this.sampleGrabber);
        }
    }
}
