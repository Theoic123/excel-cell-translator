using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelCellTranslator
{
    internal sealed class ProviderSettingsForm : Form
    {
        private readonly ComboBox primary = new ComboBox();
        private readonly ComboBox fallback = new ComboBox();
        private readonly TextBox google = new TextBox();
        private readonly TextBox deepl = new TextBox();
        private readonly CheckBox pro = new CheckBox();

        internal ProviderSettingsForm()
        {
            Text = "翻译服务";
            Font = new Font("Microsoft YaHei UI", 9F);
            ClientSize = new Size(510, 520);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = MinimizeBox = false;
            AddLabel("首选服务", 20);
            AddLabel("备用服务", 63);
            primary.SetBounds(125, 17, 355, 28);
            fallback.SetBounds(125, 60, 355, 28);
            primary.DropDownStyle = fallback.DropDownStyle = ComboBoxStyle.DropDownList;
            string[] names = { "MyMemory（无需密钥）", "Google Cloud Translation", "DeepL API" };
            primary.Items.AddRange(names);
            fallback.Items.Add("不启用自动备用");
            fallback.Items.AddRange(names);
            AddLabel("Google API key", 111);
            google.SetBounds(24, 139, 456, 27);
            google.UseSystemPasswordChar = true;
            AddLabel("DeepL API key", 184);
            deepl.SetBounds(24, 212, 456, 27);
            deepl.UseSystemPasswordChar = true;
            pro.Text = "使用 DeepL API Pro（API Free 请不要勾选）";
            pro.SetBounds(24, 250, 456, 26);
            var notice = new Label { Text = "Google / DeepL 需要官方 API 密钥，网页版账号不能代替。\r\n启用备用后，额度耗尽或网络故障时会将原文发送给备用服务。\r\n密钥在本机按 Windows 用户加密保存；服务用量可能产生费用。", Location = new Point(24, 293), Size = new Size(456, 150) };
            var save = new Button { Text = "保存服务", Location = new Point(365, 472), Size = new Size(115, 30) };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Location = new Point(260, 472), Size = new Size(90, 30) };
            save.Click += Save;
            Controls.AddRange(new Control[] { primary, fallback, google, deepl, pro, notice, save, cancel });
            AcceptButton = save; CancelButton = cancel;
            try
            {
                ProviderSettings settings = ProviderSettings.Load();
                primary.SelectedIndex = (int)settings.PrimaryProvider;
                fallback.SelectedIndex = settings.FallbackProvider.HasValue ? (int)settings.FallbackProvider.Value + 1 : 0;
                google.Text = settings.GoogleApiKey; deepl.Text = settings.DeepLApiKey;
                pro.Checked = settings.DeepLUsePro;
            }
            catch (Exception)
            {
                primary.SelectedIndex = 0; fallback.SelectedIndex = 0;
                notice.Text = "保存的服务设置无法读取，请重新填写。密钥不会在错误信息中显示。";
            }
        }

        private void AddLabel(string text, int y)
        { Controls.Add(new Label { Text = text, Location = new Point(24, y), AutoSize = true }); }

        private void Save(object sender, EventArgs e)
        {
            var settings = new ProviderSettings {
                PrimaryProvider = (TranslationProvider)primary.SelectedIndex,
                FallbackProvider = fallback.SelectedIndex == 0 ? (TranslationProvider?)null : (TranslationProvider)(fallback.SelectedIndex - 1),
                GoogleApiKey = google.Text.Trim(), DeepLApiKey = deepl.Text.Trim(), DeepLUsePro = pro.Checked
            };
            if (!settings.IsProviderConfigured(settings.PrimaryProvider) ||
                (settings.FallbackProvider.HasValue && !settings.IsProviderConfigured(settings.FallbackProvider.Value)))
            { MessageBox.Show(this, "请填写所选首选服务和备用服务的 API 密钥。", "缺少密钥"); return; }
            if (settings.FallbackProvider == settings.PrimaryProvider)
            { MessageBox.Show(this, "备用服务应与首选服务不同。", "请选择其他备用服务"); return; }
            try { settings.Save(); DialogResult = DialogResult.OK; Close(); }
            catch (Exception) { MessageBox.Show(this, "无法保存服务设置，请检查当前用户的文件访问权限。", "保存失败"); }
        }
    }
}
