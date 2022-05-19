using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using TFlex.DOCs.Model.Macros;
using TFlex.DOCs.Model.Macros.ObjectModel;

using System.Windows.Forms.DataVisualization.Charting;
using System.Drawing;

public class Macro : MacroProvider
{
    public Macro(MacroContext context)
        : base(context)
    {
    }

    public override void Run()
    {
    	
    }
    
    public void ПоказатьДиаграмму(string заголовокДиаграммы, string[,] Вывод_на_диаграмму, double total)
    {
    	using (MyChartForm dialog = new MyChartForm(заголовокДиаграммы, Вывод_на_диаграмму, total)) 
		{
			dialog.SaveImage();
			/*Не показывать изображение
			if (dialog.ShowDialog() == DialogResult.OK)
			 {}
			*/
		}
    }
 }
 

 
partial class MyChartForm : Form
{
		public MyChartForm(string заголовокДиаграммы, string[,] Вывод_на_диаграмму, double total)
		{
      		InitializeComponent(заголовокДиаграммы, Вывод_на_диаграмму, total);
		}
		
		public void SaveImage ()
		{
			/*png - формат*/
			string tempPath = System.IO.Path.GetTempPath() + "\\Chart.png";
			this.chart1.SaveImage(tempPath, ChartImageFormat.Png);
			
			/*bmp - формат
			string tempPath = System.IO.Path.GetTempPath() + "\\Chart.bmp";
			this.chart1.SaveImage(tempPath, ChartImageFormat.Bmp);
			*/
		}
		
		private string format = "{0:0.##} ({1:0.#}%)";

		/// <summary>
		/// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;


        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent(string заголовокДиаграммы, string[,] Вывод_на_диаграмму, double total)
        {
            ChartArea chartArea1 = new ChartArea();
            Legend legend1 = new Legend();
            Series series1 = new Series();
            
            int size = Вывод_на_диаграмму.GetLength(0);
            DataPoint[] dataPoint = new DataPoint[size];
            
            Title title1 = new Title();
            this.chart1 = new Chart();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            this.SuspendLayout();
            // 
            // chart1
            // 
            chartArea1.Area3DStyle.Enable3D = true;
            chartArea1.Area3DStyle.Inclination = 50;
            chartArea1.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea1);
            legend1.Font = new System.Drawing.Font("T-FLEX Type A", 16F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204)));
            legend1.IsTextAutoFit = false;
            legend1.LegendStyle = LegendStyle.Column;
            legend1.InterlacedRows = true;
            legend1.InterlacedRowsColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(230)))));
            legend1.TextWrapThreshold = 36;
            legend1.Name = "Legend1";
            legend1.Title = "Описание:";
            legend1.TitleFont = new System.Drawing.Font("T-FLEX Type A", 16F, FontStyle.Underline, GraphicsUnit.Point, ((byte)(204)));
            this.chart1.Legends.Add(legend1);
            this.chart1.Location = new System.Drawing.Point(21, 12);
            this.chart1.Name = "chart1";
            series1.ChartArea = "ChartArea1";
            series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
            series1.CustomProperties = "MinimumRelativePieSize=40, 3DLabelLineSize=30";
            series1.Legend = "Legend1";
            series1.Name = "Series1";
            for (int i = 0; i < size; i++)
            {
            	string label = Вывод_на_диаграмму[i, 0];
            	double time = Convert.ToDouble(Вывод_на_диаграмму[i, 1]);
            	double percent = time/total*100;
            	string legend = Вывод_на_диаграмму[i, 0] + " - " + Вывод_на_диаграмму[i, 2] + ": " + string.Format(format, time, percent);
            	
            	dataPoint[i] = new DataPoint(0D, time);
            	FillDataPoint(ref dataPoint[i], label, legend);
            	series1.Points.Add(dataPoint[i]);
            }
            this.chart1.Series.Add(series1);
            this.chart1.Size = new Size(1500, 1000);
            this.chart1.TabIndex = 0;
            this.chart1.Text = "chart1";
            title1.Font = new System.Drawing.Font("T-FLEX Type A", 20F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204)));
            title1.Name = "Title1";
            title1.Text = заголовокДиаграммы;
            this.chart1.Titles.Add(title1);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1530, 1020);
            this.Controls.Add(this.chart1);
            this.Name = "Form1";
            this.Text = "Экспресс-анализ";
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.ResumeLayout(false);

        }
        
        private void FillDataPoint (ref System.Windows.Forms.DataVisualization.Charting.DataPoint dataPoint, string label, string legend)
        {
        	dataPoint.Font = new System.Drawing.Font("T-FLEX Type A", 14F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204)));
            dataPoint.IsValueShownAsLabel = false;
            dataPoint.Label = label;
            dataPoint.LegendText = legend;
            dataPoint.CustomProperties = "PieLabelStyle=Outside";
        }
        
        /*Для вывода описания технологических переделов вместо кода (не используется)*/
        private string DivideLabel (string input)
        {
        	int size = 22;
        	string info = "";
        	string temp1;
    		int index_prob;
    		int index_defice;
    		int index;
        	string temp2 = input;
        	
	    	if (input.Length <= size)
	    	{
	    		index_prob = temp2.LastIndexOf(" ");
	    		index_defice = temp2.LastIndexOf("-");
	    		if (index_prob > 7 && index_prob < 15)
	    		{
	    			temp1 = temp2.Substring(0, index_prob);
	    			info += temp1 + "\r\n";
	    			info += temp2.Remove(0, index_prob + 1);
	    		}
	    		else if (index_defice > 7 && index_defice < 15)
	    		{
	    			temp1 = temp2.Substring(0, index_defice + 1);
	    			info += temp1 + "\r\n";
	    			info += temp2.Remove(0, index_defice + 1);
	    		}
	    		else
	    			info += input;
	    	}
	    	else
	    	{
	    		while (temp2.Length > size)
	    		{
	    			temp1 = temp2.Substring(0, size);
	    			index_prob = temp1.LastIndexOf(" ");
	    			index_defice = temp1.LastIndexOf("-");
	    			index = Math.Max(index_prob, index_defice);
	    			temp1 = temp1.Substring(0, index);
	    			info += temp1 + "\r\n";
	    			temp2 = temp2.Remove(0, index);
	    			if (index_prob == index)
	    				temp2 = temp2.Remove(0, 1);
	    		}
	    			
	    		info += temp2;
	    	}
	    	return info;
        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
}
