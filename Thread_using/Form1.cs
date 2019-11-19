using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Thread_using
{
    public partial class Form1 : Form
    {
        public static TextBox t;
        public Form1()
        {
            InitializeComponent();
            System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;
            t = textBox1;
        }
        public void method()
        {
            //Console.WriteLine("这是无参的静态方法");
            textBox1.Text += "创建了一个线程" + "\r\n";
        }
        class ThreadTest
        {
            public void MyThread()
            {
                Console.WriteLine("这是一个实例方法");
            }
        }
        static void Thread1(object obj)
        {
            Console.WriteLine(obj);
        }


        private void button2_Click(object sender, EventArgs e)
        {
            //创建无参数方法的托管线程
            //创建线程
            Thread thread1 = new Thread(new ThreadStart(method));
            //启动线程
            thread1.Start();
            ThreadTest test = new ThreadTest();
            //创建线程
            Thread thread2 = new Thread(new ThreadStart(test.MyThread));
            //启动线程
            thread2.Start();
            Thread thread3 = new Thread(delegate () { Console.WriteLine("我是通过匿名委托创建的线程"); });
            thread3.Start();

            //通过Lambda表达式创建
            Thread thread4 = new Thread(() => { Console.WriteLine("我是通过Lambda表达式创建的委托"); });
            thread4.Start();
            //通过ParameterizedThreadStart创建线程
            Thread thread = new Thread(new ParameterizedThreadStart(Thread1));
            //给方法传值
            thread.Start("这是一个有参数的委托");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text += "当前线程id为" + Thread.CurrentThread.ManagedThreadId.ToString("00") + "\r\n";
        }
        class ThreadSample
        {
            private readonly int _iterations;
            public int num = 0;
            public ThreadSample(int iterations)
            {
                _iterations = iterations;
            }

            public void CountNumbers()
            {

                for (int i = 0; i < _iterations; i++)
                {

                    Thread.Sleep(TimeSpan.FromSeconds(0.5));
                    t.Text += Thread.CurrentThread.Name + "计数 " + i + "\r\n";
                }

            }
            public void plus()
            {
                lock (this)
                {
                    this.num += 1;
                    t.Text += Thread.CurrentThread.Name + "计数 " + num + "\r\n";
                }

            }
            public void muPlus()
            {
                for (int i = 0; i < _iterations; i++)
                {
                    plus();
                }
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            var sampleForeground = new ThreadSample(5);

            var threadOne = new Thread(sampleForeground.CountNumbers);
            threadOne.Name = "前台线程";



            threadOne.Start();


        }

        private void button4_Click(object sender, EventArgs e)
        {
            var sampleBackground = new ThreadSample(5);
            var threadTwo = new Thread(sampleBackground.CountNumbers);
            threadTwo.IsBackground = true;
            threadTwo.Name = "后台线程";
            threadTwo.Start();
        }

        private void button5_Click(object sender, EventArgs e)
        {

            var sampleBackground = new ThreadSample(5);
            var threadTwo = new Thread(sampleBackground.CountNumbers);
            threadTwo.Name = "子线程";
            threadTwo.Start();

            threadTwo.Join();
            t.Text += "子线程计数完成\r\n";

        }

        private void button11_Click(object sender, EventArgs e)
        {
            var sampleBackground = new ThreadSample(10);
            AsyncCallback callback = p =>
            {
                t.Text += ($"到这里已经完成了。{Thread.CurrentThread.ManagedThreadId.ToString("00")}。");

            };
            //异步调用回调
            Action fn = sampleBackground.CountNumbers;
            for (int i = 0; i < 5; i++)
            {
                string name = string.Format($"btnSync_Click_{i}");
                fn.BeginInvoke(callback, null);
            }

        }


        private void button6_Click(object sender, EventArgs e)
        {

            var sampleBackground = new ThreadSample(5);

            var threadOne = new Thread(sampleBackground.muPlus);
            var threadTwo = new Thread(sampleBackground.muPlus);

            threadOne.Name = "线程1";
            threadTwo.Name = "线程2";
            threadOne.Start();
            threadTwo.Start();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            var sampleBackground = new ThreadSample(10);

            var threadOne = new Thread(sampleBackground.CountNumbers);
            threadOne.Name = "线程";
            threadOne.Start();
            textBox1.Text += "结束\r\n";
        }
        static EventWaitHandle _tollStation = new AutoResetEvent(false);//车闸默认关闭
        static void Count()
        {

            _tollStation.WaitOne();
            for (int i = 0; i < 5; i++)
            {

                Thread.Sleep(TimeSpan.FromSeconds(0.5));
                t.Text += Thread.CurrentThread.Name + "计数 " + i + "\r\n";
            }
        }
        static void Count2()
        {

            _tollStation2.WaitOne();
            for (int i = 0; i < 5; i++)
            {

                Thread.Sleep(TimeSpan.FromSeconds(0.5));
                t.Text += Thread.CurrentThread.Name + "计数 " + i + "\r\n";
            }
        }
        private void button8_Click(object sender, EventArgs e)
        {
            var threadOne = new Thread(Count);
            threadOne.Name = "线程";
            threadOne.Start();
            Thread.Sleep(TimeSpan.FromSeconds(1));
            _tollStation.Set();
        }
        static EventWaitHandle _tollStation2 = new ManualResetEvent(false);//改为ManualResetEvent,车闸默认关闭
        private void button9_Click(object sender, EventArgs e)
        {


            var threadOne = new Thread(Count2);
            var threadTwo = new Thread(Count2);

            threadOne.Name = "线程1";
            threadTwo.Name = "线程2";
            threadOne.Start();
            threadTwo.Start();
            Thread.Sleep(TimeSpan.FromSeconds(1));
            _tollStation2.Set();
        }
        static EventWaitHandle _tollStation3 = new AutoResetEvent(false);
        class Test
        {
            public int num = 0;
            public void product()
            {
                for (int i = 0; i < 10; i++)
                {
                    lock (this)
                    {
                        num += 1;
                        if (num ==1 )
                        {
                            _tollStation3.Set();
                        }
                        lock (t)
                        {
                            t.Text += "生产者生产后，总数为" + num + "\r\n";
                        }
                    }


                }
            }
            public void consume()
            {
                for (int i = 0; i < 10; i++)
                {
                    lock (this)
                    {
                        if (num <= 0)
                        {
                            _tollStation3.WaitOne();
                        }
                        num -= 1;
                        lock (t)
                        {
                            t.Text += "消费者消费后，总数为" + num + "\r\n";
                        }
                    }
                }
            }

        }

        private void button10_Click(object sender, EventArgs e)
        {
            Test t = new Test();
            var threadOne = new Thread(t.product);
            var threadTwo = new Thread(t.consume);

            threadOne.Name = "线程1";
            threadTwo.Name = "线程2";
            threadOne.Start();
            threadTwo.Start();
        }
    }
}
