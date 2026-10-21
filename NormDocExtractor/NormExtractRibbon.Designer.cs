namespace NormDocExtractor
{
  partial class NormExtractRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
  {
    /// <summary>
    /// Обязательная переменная конструктора.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    public NormExtractRibbon()
        : base(Globals.Factory.GetRibbonFactory())
    {
      InitializeComponent();
    }

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

    #region Код, автоматически созданный конструктором компонентов

    /// <summary>
    /// Требуемый метод для поддержки конструктора — не изменяйте 
    /// содержимое этого метода с помощью редактора кода.
    /// </summary>
    private void InitializeComponent()
    {
      this.tab1 = this.Factory.CreateRibbonTab();
      this.group1 = this.Factory.CreateRibbonGroup();
      this.btnExtractDocs = this.Factory.CreateRibbonButton();
      this.tab1.SuspendLayout();
      this.group1.SuspendLayout();
      this.SuspendLayout();
      // 
      // tab1
      // 
      this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
      this.tab1.ControlId.OfficeId = "TabHome";
      this.tab1.Groups.Add(this.group1);
      this.tab1.Label = "TabHome";
      this.tab1.Name = "tab1";
      // 
      // group1
      // 
      this.group1.Items.Add(this.btnExtractDocs);
      this.group1.Label = "Экстрактор";
      this.group1.Name = "group1";
      // 
      // btnExtractDocs
      // 
      this.btnExtractDocs.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
      this.btnExtractDocs.Image = global::NormDocExtractor.Properties.Resources.icons8_извлечение_данных_64;
      this.btnExtractDocs.Label = "Извлечь нормативы";
      this.btnExtractDocs.Name = "btnExtractDocs";
      this.btnExtractDocs.ShowImage = true;
      this.btnExtractDocs.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExtractDocs_Click);
      // 
      // NormExtractRibbon
      // 
      this.Name = "NormExtractRibbon";
      this.RibbonType = "Microsoft.Word.Document";
      this.Tabs.Add(this.tab1);
      this.tab1.ResumeLayout(false);
      this.tab1.PerformLayout();
      this.group1.ResumeLayout(false);
      this.group1.PerformLayout();
      this.ResumeLayout(false);

    }

    #endregion

    internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
    internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
    internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExtractDocs;
  }

  partial class ThisRibbonCollection
  {
    internal NormExtractRibbon MyRibbon
    {
      get { return this.GetRibbon<NormExtractRibbon>(); }
    }
  }
}
