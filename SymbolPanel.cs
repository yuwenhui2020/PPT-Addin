// SymbolPanel.cs
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace PowerPointAddIn
{
    public partial class SymbolPanel : UserControl
    {
        // 符号分类集合
        private Dictionary<string, List<SymbolItem>> _symbolCategories = new Dictionary<string, List<SymbolItem>>();

        // 当前显示的符号集合
        private List<SymbolItem> _currentSymbols = new List<SymbolItem>();

        // 事件：当用户请求添加符号到幻灯片
        public event Action<string, Font> AddSymbolRequested;

        public SymbolPanel()
        {
            InitializeComponent();
            LoadSymbolCategories();

            // 初始化分类下拉框
            cmbCategory.Items.AddRange(_symbolCategories.Keys.ToArray());
            //默认选择第一项
            if (cmbCategory.Items.Count > 0)
            {
                cmbCategory.SelectedIndex = 0;
            }
        }
        public void RefreshSymbols()
        {
            if (cmbCategory.Items.Count > 0 && cmbCategory.SelectedIndex < 0)
            {
                cmbCategory.SelectedIndex = 0;
            }
        }
        // 加载符号分类
        private void LoadSymbolCategories()
        {
            string[] offSymbols = { "α", "β", "Δ", "θ", "ʃ", "λ", "θ", "ρ", "ψ", "υ", "μ", "σ", "φ", "ω", "π", "ν", "τ", "ξ", "η", "δ", "←", "↓", "↑", "→", "∓", "±", "≠", "∝", "∞", "•", "≈", "≌", "≡", "‰", "‱", "∫", "∬", "∯", "∭", "∰", "⊙", "△", "▽", "∑", "㏑", "㏒", "∵", "∴", "™", "@", "®", "©", "℡", "☪", "✔", "✓", "✖", "✕", "✘", "✗", "☐", "☑", "☒", "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "℃", "℉", "¥", "$", "€", "£", "Ұ" };
            _symbolCategories.Add("常用符号", offSymbols.Select(s => new SymbolItem(s, "")).ToList());

            var GreeceSymbols = new List<SymbolItem>
            {
                new SymbolItem("Α","alpha"),
                new SymbolItem("α","alpha"),
                new SymbolItem("Β","beta"),
                new SymbolItem("β","beta"),
                new SymbolItem("Γ","gamma"),
                new SymbolItem("γ","gamma"),
                new SymbolItem("Δ","delta"),
                new SymbolItem("δ","delta"),
                new SymbolItem("Ε","epsilon"),
                new SymbolItem("ε","epsilon"),
                new SymbolItem("Ζ","zeta"),
                new SymbolItem("ζ","zeta"),
                new SymbolItem("Η","eta"),
                new SymbolItem("η","eta"),
                new SymbolItem("Θ","theta"),
                new SymbolItem("θ","theta"),
                new SymbolItem("Ι","iota"),
                new SymbolItem("ι","iota"),
                new SymbolItem("℩","iota"),
                new SymbolItem("Κ","kappa"),
                new SymbolItem("κ","kappa"),
                new SymbolItem("∧","lambda"),
                new SymbolItem("λ","lambda"),
                new SymbolItem("Μ","mu"),
                new SymbolItem("μ","mu"),
                new SymbolItem("Ν","nu"),
                new SymbolItem("ν","nu"),
                new SymbolItem("Ξ","xi"),
                new SymbolItem("ξ","xi"),
                new SymbolItem("Ο","omicron"),
                new SymbolItem("ο","omicron"),
                new SymbolItem("∏","pi"),
                new SymbolItem("π","pi"),
                new SymbolItem("Ρ","rho"),
                new SymbolItem("ρ","rho"),
                new SymbolItem("∑","sigma"),
                new SymbolItem("σ","sigma"),
                new SymbolItem("ς","sigma"),
                new SymbolItem("Τ","tau"),
                new SymbolItem("τ","tau"),
                new SymbolItem("Υ","upsilon"),
                new SymbolItem("υ","upsilon"),
                new SymbolItem("Φ","phi"),
                new SymbolItem("φ","phi"),
                new SymbolItem("Χ","chi"),
                new SymbolItem("χ","chi"),
                new SymbolItem("Ψ","psi"),
                new SymbolItem("ψ","psi"),
                new SymbolItem("Ω","omega"),
                new SymbolItem("ω","omega")
            };
            _symbolCategories.Add("希腊字母", GreeceSymbols);

            string[] mathSymbols = { "+", "-", "×", "÷", "±", "/", "=", "≌", "∽", "≦", "≧", "≒", "﹤", "﹥", "∷", "∶", "∫", "∮", "∝", "∞", "°", "∠", "∟", "⌒", "⊙", "∧", "∨", "∑", "∏", "∪", "∩", "∈", "⊆", "⊇", "∀", "∃", "∄", "∵", "∴", "⊥", "∥", "∠", "⌒", "⊙", "√", "∟", "⊿", "㏒", "㏑", "%", "‰", "⅟", "½", "⅓", "⅕", "⅙", "⅛", "⅔", "⅖", "⅚", "⅜", "¾", "⅗", "⅝", "⅞", "⅘", "〜", "≈", "≃", "≄", "≅", "≆", "≇", "≈", "≉", "≊", "≋", "≌", "≍", "≎", "≏", "≐", "≑", "≒", "≓", "≔", "≕", "≖", "≗", "≘", "≙", "≚", "≛", "≜", "≝", "≞", "≟", "≠", "≡", "≢", "≣", "≤", "≥", "≪", "≫", "≦", "≧", "≨", "≩", "<", ">", "≮", "≯", "⊰", "⊱", "⋛", "⋚", "⇒", "∫", "∬", "∭", "∮", "∯", "∰", "∱", "∲", "∳", "%", "‰", "‱", "ø", "Ø", "Φ", "π", "Δ" };
            _symbolCategories.Add("数学符号", mathSymbols.Select(s => new SymbolItem(s, "")).ToList());

            string[] lineSymbols = { "─", "━", "│", "┃", "┄", "┅", "┆", "┇", "┈", "┉", "┊", "┌", "┍", "┎", "┏", "┐", "┑", "┒", "┓", "└", "┖", "┕", "┗", "┘", "┙", "┚", "┛", "├", "┝", "┞", "┟", "┠", "┡", "┢", "┣", "┤", "┥", "┦", "┧", "┨", "┩", "┪", "┫", "┬", "┭", "┮", "┯", "┰", "┱", "┲", "┳", "┴", "┵", "┶", "┷", "┸", "┹", "┺", "┻", "┼", "┽", "┾", "┿", "╀", "╁", "╂", "╃", "╄", "╅", "╆", "╇", "╈", "╉", "╊", "╋", "═", "║", "╒", "╓", "╔", "╕", "╖", "╗", "╘", "╙", "╚", "╛", "╜", "╝", "╞", "╟", "╠", "╡", "╢", "╣", "╤", "╥", "╦", "╧", "╨", "╩", "╪", "╫", "╬" };
            _symbolCategories.Add("制表符", lineSymbols.Select(s => new SymbolItem(s, "")).ToList());
            string[] numberSymbols = { "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫", "⑬", "⑭", "⑮", "⑯", "⑰", "⑱", "⑲", "⑳", "⓪", "❶", "❷", "❸", "❹", "❺", "❻", "❼", "❽", "❾", "❿", "⓫", "⓬", "⓭", "⓮", "⓯", "⓰", "⓱", "⓲", "⓳", "⓴", "㊀", "㊁", "㊂", "㊃", "㊄", "㊅", "㊆", "㊇", "㊈", "㊉", "㈠", "㈡", "㈢", "㈣", "㈤", "㈥", "㈦", "㈧", "㈨", "㈩", "⑴", "⑵", "⑶", "⑷", "⑸", "⑹", "⑺", "⑻", "⑼", "⑽", "⑾", "⑿", "⒀", "⒁", "⒂", "⒃", "⒄", "⒅", "⒆", "⒇", "⒈", "⒉", "⒊", "⒋", "⒌", "⒍", "⒎", "⒏", "⒐", "⒑", "⒒", "⒓", "⒔", "⒕", "⒖", "⒗", "⒘", "⒙", "⒚", "⒛", "Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ", "Ⅸ", "Ⅹ", "Ⅺ", "Ⅻ", "ⅰ", "ⅱ", "ⅲ", "ⅳ", "ⅴ", "ⅵ", "ⅶ", "ⅷ", "ⅸ", "ⅹ", "Ⓐ", "Ⓑ", "Ⓒ", "Ⓓ", "Ⓔ", "Ⓕ", "Ⓖ", "Ⓗ", "Ⓘ", "Ⓙ", "Ⓚ", "Ⓛ", "Ⓜ", "Ⓝ", "Ⓞ", "Ⓟ", "Ⓠ", "Ⓡ", "Ⓢ", "Ⓣ", "Ⓤ", "Ⓥ", "Ⓦ", "Ⓧ", "Ⓨ", "Ⓩ", "ⓐ", "ⓑ", "ⓒ", "ⓓ", "ⓔ", "ⓕ", "ⓖ", "ⓗ", "ⓘ", "ⓙ", "ⓚ", "ⓛ", "ⓜ", "ⓝ", "ⓞ", "ⓟ", "ⓠ", "ⓡ", "ⓢ", "ⓣ", "ⓤ", "ⓥ", "ⓦ", "ⓧ", "ⓨ", "ⓩ", "⒜", "⒝", "⒞", "⒟", "⒠", "⒡", "⒢", "⒣", "⒤", "⒥", "⒦", "⒧", "⒨", "⒩", "⒪", "⒫", "⒬", "⒭", "⒮", "⒯", "⒰", "⒱", "⒲", "⒳", "⒴", "⒵" };
            _symbolCategories.Add("序号符号", numberSymbols.Select(s => new SymbolItem(s, "")).ToList());
        }
        //
        // 分类选择变化时
        private void cmbCategory_SelectedIndexChanged(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();

            if (cmbCategory.SelectedItem != null)
            {
                var category = cmbCategory.SelectedItem.ToString();
                _currentSymbols = _symbolCategories[category];

                foreach (var symbol in _currentSymbols)
                {
                    var btn = new Button
                    {
                        Text = symbol.Symbol,
                        Tag = symbol,
                        Width = 40,
                        Height = 40,
                        Font = new Font("Arial Unicode MS", 14),
                        FlatStyle = FlatStyle.Flat
                    };
                    btn.Click += SymbolButton_Click;

                    var toolTip = new ToolTip();
                    toolTip.SetToolTip(btn, symbol.Description);

                    flowLayoutPanel1.Controls.Add(btn);
                }
            }
        }

        private void SymbolButton_Click(object sender, EventArgs e)
        {
            var button = sender as Button;
            var symbolItem = button.Tag as SymbolItem;

            AddSymbolRequested?.Invoke(symbolItem.Symbol, button.Font);
        }

        // 符号项类
        private class SymbolItem
        {
            public string Symbol { get; set; }
            public string Description { get; set; }

            public SymbolItem(string symbol, string description)
            {
                Symbol = symbol;
                Description = description;
            }
        }

        private void AddBulkSymbols(string category, string symbols)
        {
            if (!_symbolCategories.ContainsKey(category))
            {
                _symbolCategories[category] = new List<SymbolItem>();
            }

            foreach (char symbol in symbols)
            {
                _symbolCategories[category].Add(new SymbolItem(symbol.ToString(), ""));
            }
        }

        private void AddSymbolWithDescription(string category, string symbol, string description)
        {
            if (!_symbolCategories.ContainsKey(category))
            {
                _symbolCategories[category] = new List<SymbolItem>();
            }

            _symbolCategories[category].Add(new SymbolItem(symbol, description));
        }
    }
}