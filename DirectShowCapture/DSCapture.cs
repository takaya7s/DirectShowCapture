using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using DirectShowLib;

namespace DirectShowCapture
{
    public class DSCapture : IDisposable, ISampleGrabberCB
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

        private readonly object _bufLock = new object();

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

        public DSCapture(DsDevice device, IntPtr Handle, Size size) : base()
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
            lock (_bufLock)
            {
                MoveMemory(this.buffer, pBuffer, bufferLength);
            }
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

        private Bitmap _reusableBmp;
        private BitmapData _reusableLock;

        /// <summary>
        /// Bitmapに変換します
        /// </summary>
        /// <returns></returns>
        public Bitmap Capture()
        {
            if (buffer == IntPtr.Zero) return null;

            if (_reusableBmp == null ||
                _reusableBmp.Width != width || _reusableBmp.Height != height)
            {
                _reusableBmp?.Dispose();
                _reusableBmp = new Bitmap(width, height, PixelFormat.Format24bppRgb);
            }

            Rectangle rect = new Rectangle(0, 0, width, height);
            lock (_bufLock)
            {
                var data = _reusableBmp.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);
                try
                {
                    // ここで上下反転しつつコピー（行逆順）
                    int srcStride = stride;
                    int dstStride = data.Stride;
                    unsafe
                    {
                        byte* dst0 = (byte*)data.Scan0;
                        byte* src0 = (byte*)buffer;
                        for (int y = 0; y < height; y++)
                        {
                            // bottom-up → top-down
                            byte* src = src0 + (height - 1 - y) * srcStride;
                            byte* dst = dst0 + y * dstStride;
                            Buffer.MemoryCopy(src, dst, dstStride, Math.Min(dstStride, srcStride));
                        }
                    }
                }
                finally
                {
                    _reusableBmp.UnlockBits(data);
                }
            }
            return (Bitmap)_reusableBmp.Clone();
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
        /// 
        /// </summary>
        /// <param name="n"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public Bitmap CaptureDownscaleByBlock(int n)
        {
            if (buffer == IntPtr.Zero) return null;
            if (n <= 0) throw new ArgumentOutOfRangeException(nameof(n));
            if (width % n != 0 || height % n != 0)
                throw new ArgumentException("元画像が n で割り切れません。");

            int dw = width / n;
            int dh = height / n;
            var dst = new Bitmap(dw, dh, PixelFormat.Format24bppRgb);
            var rect = new Rectangle(0, 0, dw, dh);
            var dstData = dst.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);

            try
            {
                lock (_bufLock)
                {
                    unsafe
                    {
                        byte* src0 = (byte*)buffer; // bottom-up
                        byte* d0 = (byte*)dstData.Scan0;
                        int dStride = dstData.Stride;

                        long area = (long)n * n;

                        for (int oy = 0; oy < dh; oy++)
                        {
                            byte* dp = d0 + oy * dStride;
                            int syTop = height - (oy * n) - 1; // bottom-up → top-down
                                                               // ブロックの上端から下へ n 行
                            for (int ox = 0; ox < dw; ox++)
                            {
                                long sumR = 0, sumG = 0, sumB = 0;

                                int sxLeft = ox * n;

                                for (int ky = 0; ky < n; ky++)
                                {
                                    int srcY = syTop - ky; // 実メモリ行
                                    byte* sp = src0 + srcY * stride + sxLeft * 3;

                                    for (int kx = 0; kx < n; kx++)
                                    {
                                        sumB += sp[0];
                                        sumG += sp[1];
                                        sumR += sp[2];
                                        sp += 3;
                                    }
                                }

                                dp[2] = (byte)((sumR + (area >> 1)) / area);
                                dp[1] = (byte)((sumG + (area >> 1)) / area);
                                dp[0] = (byte)((sumB + (area >> 1)) / area);
                                dp += 3;
                            }
                        }
                    }
                }
            }
            finally
            {
                dst.UnlockBits(dstData);
            }

            return dst;
        }

        public Bitmap CaptureDownscaleNearest(int dstW, int dstH)
        {
            if (buffer == IntPtr.Zero) return null;
            if (dstW <= 0 || dstH <= 0) return null;

            var dst = new Bitmap(dstW, dstH, PixelFormat.Format24bppRgb);
            var rect = new Rectangle(0, 0, dstW, dstH);
            var dstData = dst.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);

            try
            {
                double scaleX = (double)width / dstW;
                double scaleY = (double)height / dstH;

                lock (_bufLock)
                {
                    unsafe
                    {
                        byte* src0 = (byte*)buffer;       // bottom-up
                        byte* d0 = (byte*)dstData.Scan0;
                        int dStride = dstData.Stride;

                        for (int dy = 0; dy < dstH; dy++)
                        {
                            // 中心合わせの写像（+0.5 → -0.5）
                            double syf = (dy + 0.5) * scaleY - 0.5;
                            int sy = (int)Math.Round(syf);
                            if (sy < 0) sy = 0; else if (sy >= height) sy = height - 1;

                            int memY = height - 1 - sy; // bottom-up → top-down
                            byte* srow = src0 + memY * stride;

                            byte* dp = d0 + dy * dStride;

                            for (int dx = 0; dx < dstW; dx++)
                            {
                                double sxf = (dx + 0.5) * scaleX - 0.5;
                                int sx = (int)Math.Round(sxf);
                                if (sx < 0) sx = 0; else if (sx >= width) sx = width - 1;

                                byte* sp = srow + sx * 3;
                                dp[0] = sp[0]; // B
                                dp[1] = sp[1]; // G
                                dp[2] = sp[2]; // R
                                dp += 3;
                            }
                        }
                    }
                }
            }
            finally
            {
                dst.UnlockBits(dstData);
            }
            return dst;
        }

        public Bitmap CaptureDownscaleBilinear(int dstW, int dstH)
        {
            if (buffer == IntPtr.Zero) return null;
            if (dstW <= 0 || dstH <= 0) return null;

            var dst = new Bitmap(dstW, dstH, PixelFormat.Format24bppRgb);
            var rect = new Rectangle(0, 0, dstW, dstH);
            var dstData = dst.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);

            try
            {
                double scaleX = (double)width / dstW;
                double scaleY = (double)height / dstH;

                lock (_bufLock)
                {
                    unsafe
                    {
                        byte* src0 = (byte*)buffer;       // bottom-up
                        byte* d0 = (byte*)dstData.Scan0;
                        int dStride = dstData.Stride;

                        for (int dy = 0; dy < dstH; dy++)
                        {
                            double sy = (dy + 0.5) * scaleY - 0.5;
                            int y0 = (int)Math.Floor(sy);
                            double wy = sy - y0;
                            int y1 = y0 + 1;

                            if (y0 < 0) { y0 = 0; wy = 0; y1 = 0; }
                            if (y1 >= height) { y1 = height - 1; if (y0 > y1) y0 = y1; }

                            int memY0 = height - 1 - y0;
                            int memY1 = height - 1 - y1;

                            byte* row0 = src0 + memY0 * stride;
                            byte* row1 = src0 + memY1 * stride;

                            byte* dp = d0 + dy * dStride;

                            for (int dx = 0; dx < dstW; dx++)
                            {
                                double sx = (dx + 0.5) * scaleX - 0.5;
                                int x0 = (int)Math.Floor(sx);
                                double wx = sx - x0;
                                int x1 = x0 + 1;

                                if (x0 < 0) { x0 = 0; wx = 0; x1 = 0; }
                                if (x1 >= width) { x1 = width - 1; if (x0 > x1) x0 = x1; }

                                byte* p00 = row0 + x0 * 3; // (x0,y0)
                                byte* p10 = row0 + x1 * 3; // (x1,y0)
                                byte* p01 = row1 + x0 * 3; // (x0,y1)
                                byte* p11 = row1 + x1 * 3; // (x1,y1)

                                double k00 = (1.0 - wx) * (1.0 - wy);
                                double k10 = wx * (1.0 - wy);
                                double k01 = (1.0 - wx) * wy;
                                double k11 = wx * wy;

                                // B
                                double b = p00[0] * k00 + p10[0] * k10 + p01[0] * k01 + p11[0] * k11;
                                // G
                                double g = p00[1] * k00 + p10[1] * k10 + p01[1] * k01 + p11[1] * k11;
                                // R
                                double r = p00[2] * k00 + p10[2] * k10 + p01[2] * k01 + p11[2] * k11;

                                // 0..255 に丸め
                                int bi = (int)(b + 0.5); if (bi < 0) bi = 0; else if (bi > 255) bi = 255;
                                int gi = (int)(g + 0.5); if (gi < 0) gi = 0; else if (gi > 255) gi = 255;
                                int ri = (int)(r + 0.5); if (ri < 0) ri = 0; else if (ri > 255) ri = 255;

                                dp[0] = (byte)bi;
                                dp[1] = (byte)gi;
                                dp[2] = (byte)ri;
                                dp += 3;
                            }
                        }
                    }
                }
            }
            finally
            {
                dst.UnlockBits(dstData);
            }
            return dst;
        }

        public Bitmap CaptureDownscaleBicubic(int dstW, int dstH)
        {
            if (buffer == IntPtr.Zero) return null;
            if (dstW <= 0 || dstH <= 0) return null;

            var dst = new Bitmap(dstW, dstH, PixelFormat.Format24bppRgb);
            var rect = new Rectangle(0, 0, dstW, dstH);
            var dstData = dst.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format24bppRgb);

            try
            {
                double scaleX = (double)width / dstW;
                double scaleY = (double)height / dstH;

                lock (_bufLock)
                {
                    unsafe
                    {
                        byte* src0 = (byte*)buffer;   // bottom-up
                        byte* d0 = (byte*)dstData.Scan0;
                        int dStride = dstData.Stride;

                        for (int dy = 0; dy < dstH; dy++)
                        {
                            // 出力ピクセル中心を入力座標へ
                            double sy = (dy + 0.5) * scaleY - 0.5;
                            int yBase = (int)Math.Floor(sy);
                            double fy = sy - yBase; // 0..1

                            // 参照Y（4本）と係数
                            int[] ys = new int[4];
                            ys[0] = Clamp(yBase - 1, 0, height - 1);
                            ys[1] = Clamp(yBase + 0, 0, height - 1);
                            ys[2] = Clamp(yBase + 1, 0, height - 1);
                            ys[3] = Clamp(yBase + 2, 0, height - 1);

                            double[] wy = new double[4];
                            wy[0] = MitchellCubic(1.0 + fy);
                            wy[1] = MitchellCubic(fy);
                            wy[2] = MitchellCubic(1.0 - fy);
                            wy[3] = MitchellCubic(2.0 - fy);

                            // bottom-up → メモリ上のYへ
                            byte* row0 = src0 + (height - 1 - ys[0]) * stride;
                            byte* row1 = src0 + (height - 1 - ys[1]) * stride;
                            byte* row2 = src0 + (height - 1 - ys[2]) * stride;
                            byte* row3 = src0 + (height - 1 - ys[3]) * stride;

                            byte* dp = d0 + dy * dStride;

                            for (int dx = 0; dx < dstW; dx++)
                            {
                                double sx = (dx + 0.5) * scaleX - 0.5;
                                int xBase = (int)Math.Floor(sx);
                                double fx = sx - xBase;

                                int x0 = Clamp(xBase - 1, 0, width - 1);
                                int x1 = Clamp(xBase + 0, 0, width - 1);
                                int x2 = Clamp(xBase + 1, 0, width - 1);
                                int x3 = Clamp(xBase + 2, 0, width - 1);

                                double w0 = MitchellCubic(1.0 + fx);
                                double w1 = MitchellCubic(fx);
                                double w2 = MitchellCubic(1.0 - fx);
                                double w3 = MitchellCubic(2.0 - fx);

                                // 4x4 の分離畳み込み（まず横でまとめ、その縦線形結合でもOKだが、ここでは直積で素直に）
                                double b = 0, g = 0, r = 0;

                                // y0
                                {
                                    byte* p0 = row0 + x0 * 3;
                                    byte* p1 = row0 + x1 * 3;
                                    byte* p2 = row0 + x2 * 3;
                                    byte* p3 = row0 + x3 * 3;

                                    double bx = p0[0] * w0 + p1[0] * w1 + p2[0] * w2 + p3[0] * w3;
                                    double gx = p0[1] * w0 + p1[1] * w1 + p2[1] * w2 + p3[1] * w3;
                                    double rx = p0[2] * w0 + p1[2] * w1 + p2[2] * w2 + p3[2] * w3;

                                    b += bx * wy[0]; g += gx * wy[0]; r += rx * wy[0];
                                }
                                // y1
                                {
                                    byte* p0 = row1 + x0 * 3;
                                    byte* p1 = row1 + x1 * 3;
                                    byte* p2 = row1 + x2 * 3;
                                    byte* p3 = row1 + x3 * 3;

                                    double bx = p0[0] * w0 + p1[0] * w1 + p2[0] * w2 + p3[0] * w3;
                                    double gx = p0[1] * w0 + p1[1] * w1 + p2[1] * w2 + p3[1] * w3;
                                    double rx = p0[2] * w0 + p1[2] * w1 + p2[2] * w2 + p3[2] * w3;

                                    b += bx * wy[1]; g += gx * wy[1]; r += rx * wy[1];
                                }
                                // y2
                                {
                                    byte* p0 = row2 + x0 * 3;
                                    byte* p1 = row2 + x1 * 3;
                                    byte* p2 = row2 + x2 * 3;
                                    byte* p3 = row2 + x3 * 3;

                                    double bx = p0[0] * w0 + p1[0] * w1 + p2[0] * w2 + p3[0] * w3;
                                    double gx = p0[1] * w0 + p1[1] * w1 + p2[1] * w2 + p3[1] * w3;
                                    double rx = p0[2] * w0 + p1[2] * w1 + p2[2] * w2 + p3[2] * w3;

                                    b += bx * wy[2]; g += gx * wy[2]; r += rx * wy[2];
                                }
                                // y3
                                {
                                    byte* p0 = row3 + x0 * 3;
                                    byte* p1 = row3 + x1 * 3;
                                    byte* p2 = row3 + x2 * 3;
                                    byte* p3 = row3 + x3 * 3;

                                    double bx = p0[0] * w0 + p1[0] * w1 + p2[0] * w2 + p3[0] * w3;
                                    double gx = p0[1] * w0 + p1[1] * w1 + p2[1] * w2 + p3[1] * w3;
                                    double rx = p0[2] * w0 + p1[2] * w1 + p2[2] * w2 + p3[2] * w3;

                                    b += bx * wy[3]; g += gx * wy[3]; r += rx * wy[3];
                                }

                                // 0..255 に丸め
                                int bi = (int)Math.Round(b); if (bi < 0) bi = 0; else if (bi > 255) bi = 255;
                                int gi = (int)Math.Round(g); if (gi < 0) gi = 0; else if (gi > 255) gi = 255;
                                int ri = (int)Math.Round(r); if (ri < 0) ri = 0; else if (ri > 255) ri = 255;

                                dp[0] = (byte)bi;
                                dp[1] = (byte)gi;
                                dp[2] = (byte)ri;
                                dp += 3;
                            }
                        }
                    }
                }
            }
            finally
            {
                dst.UnlockBits(dstData);
            }

            return dst;

            // --- ローカル関数 ---
            int Clamp(int v, int lo, int hi) => v < lo ? lo : (v > hi ? hi : v);

            // Mitchell-Netravali（B=1/3, C=1/3）
            double MitchellCubic(double x)
            {
                const double B = 1.0 / 3.0;
                const double C = 1.0 / 3.0;

                x = Math.Abs(x);
                if (x < 1.0)
                {
                    return ((12 - 9 * B - 6 * C) * x * x * x
                          + (-18 + 12 * B + 6 * C) * x * x
                          + (6 - 2 * B)) / 6.0;
                }
                else if (x < 2.0)
                {
                    return ((-B - 6 * C) * x * x * x
                          + (6 * B + 30 * C) * x * x
                          + (-12 * B - 48 * C) * x
                          + (8 * B + 24 * C)) / 6.0;
                }
                else
                {
                    return 0.0;
                }
            }
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
        }/// <summary>
         /// ブロック平均で 1/n に縮小した「上下正しい」RGB24 の byte[] を返す。
         /// out: 縮小後の幅・高さ・ストライド（4バイト境界）
         /// 例：1920x1080, n=60 → 32x18, stride=Align4(32*3)=96
         /// </summary>
        public byte[] GetBufferDownscaledByBlock(int n, out int outWidth, out int outHeight)
        {
            if (buffer == IntPtr.Zero) { outWidth = outHeight = 0; return null; }
            if (n <= 0) throw new ArgumentOutOfRangeException(nameof(n));
            if (width % n != 0 || height % n != 0)
                throw new ArgumentException("元画像の幅/高さが n で割り切れません。");

            int dw = width / n;
            int dh = height / n;
            int dStride = Align4(dw * 3);
            var dst = new byte[dStride * dh];

            lock (_bufLock)
            {
                unsafe
                {
                    byte* src0 = (byte*)buffer; // bottom-up
                    long area = (long)n * n;

                    fixed (byte* d0 = dst)
                    {
                        for (int oy = 0; oy < dh; oy++)
                        {
                            byte* dp = d0 + oy * dStride;

                            for (int ox = 0; ox < dw; ox++)
                            {
                                long sumR = 0, sumG = 0, sumB = 0;

                                // n×n ブロックを積算（top-down で読む）
                                for (int ky = 0; ky < n; ky++)
                                {
                                    int srcY = height - 1 - (oy * n + ky); // bottom-up → top-down
                                    byte* sp = src0 + srcY * stride + (ox * n) * 3;

                                    for (int kx = 0; kx < n; kx++)
                                    {
                                        sumB += sp[0];
                                        sumG += sp[1];
                                        sumR += sp[2];
                                        sp += 3;
                                    }
                                }

                                // 丸め付きの平均
                                int B = (int)((sumB + (area >> 1)) / area);
                                int G = (int)((sumG + (area >> 1)) / area);
                                int R = (int)((sumR + (area >> 1)) / area);

                                // 書き込み（RGB24）
                                dp[0] = (byte)B;
                                dp[1] = (byte)G;
                                dp[2] = (byte)R;
                                dp += 3;
                            }
                        }
                    }
                }
            }

            outWidth = dw;
            outHeight = dh;
            return dst;

            // --- ローカル関数 ---
            // 4バイト境界に揃える
            int Align4(int x) => (x + 3) & ~3;
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
