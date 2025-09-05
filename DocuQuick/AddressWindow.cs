using ComponentFactory.Krypton.Ribbon;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Security.Credentials.UI;


namespace Document_Maker
{
    public partial class AddressWindow : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
        public AddressWindow()
        {
            InitializeComponent();
        }


        public string PageConfirmationResult { get; set; }

        public AddressWindow(string AddressOrCaller) : this()
        {
            PageConfirmationResult = AddressOrCaller;
        }


        public void AddressWindow_Load(object sender, EventArgs e)
        {
            if(Properties.Settings.Default.ToOrCaller == 0)
            {
                kryptonCheckBox10.Visible = false;
                kryptonCheckBox11.Visible = false;
                kryptonCheckBox12.Visible = false;
                kryptonCheckBox13.Visible = false;
            }
            else if (Properties.Settings.Default.ToOrCaller ==1)
            {
                kryptonCheckBox10.Visible = true;
                kryptonCheckBox11.Visible = true;
                kryptonCheckBox12.Visible = true;
                kryptonCheckBox13.Visible = true;
            }

            if (PageConfirmationResult.Contains("wizardPage3") == true)
            {
                this.Text = "相手先の連絡先を選択...";
                kryptonHeaderGroup1.ValuesPrimary.Heading = "相手先の連絡先";
            }
            else if (PageConfirmationResult.Contains("wizardPage4") == true)
            {
                this.Text = "発信者の連絡先を選択...";
                kryptonHeaderGroup1.ValuesPrimary.Heading = "発信者の連絡先";
            }

            if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Blue;
            }
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Silver;
            }
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2007Black;

            }
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Blue;
            }
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Silver;
            }
            else if (this.PaletteMode == ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black)
            {
                kryptonPalette1.BasePaletteMode = ComponentFactory.Krypton.Toolkit.PaletteMode.Office2010Black;
            }
        }

        //宛先の情報の読み込み
        public string AdCompanyName { get; set; }
        public string AdTitle { get; set; }
        public string AdName { get; set; }
        //発信者の情報の読み込み
        private void AddressWindow_Shown(object sender, EventArgs e)
        {
            ContactsAuth();
        }

        public void ContactUpDate()
        {
            kryptonListBox2.Items.Clear();
            //連絡先取得処理
            // Outlookアプリケーションを初期化
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

            // 連絡先フォルダを取得
            MAPIFolder contactsFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

            // 連絡先アイテムを取得
            Items contactItems = contactsFolder.Items;

            // 連絡先をループで表示
            foreach (object item in contactItems)
            {
                if (item is ContactItem contact)
                {
                    kryptonListBox2.Items.Add(contact.FullName);
                }
            }
        }

        async System.Threading.Tasks.Task ContactsAuth()
        {
            this.Enabled = false;
            var availableWindowsHello = await UserConsentVerifier.CheckAvailabilityAsync();
            if (availableWindowsHello != UserConsentVerifierAvailability.Available)
            {
                MessageBox.Show("このデバイスまたはOSでは Windows Hello をサポートしていないか設定されていない場合があります。", "認証失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                this.Close();
            }
            else
            {
                var result = await UserConsentVerifier.RequestVerificationAsync("Microsoft Outlook の連絡先を取得・使用するには認証してください。");

                if (result == UserConsentVerificationResult.Verified)
                {
                    this.Enabled = true;

                    //連絡先取得処理
                    // Outlookアプリケーションを初期化
                    Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
                    NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

                    // 連絡先フォルダを取得
                    MAPIFolder contactsFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderContacts);

                    // 連絡先アイテムを取得
                    Items contactItems = contactsFolder.Items;

                    // 連絡先をループで表示
                    foreach (object item in contactItems)
                    {
                        if (item is ContactItem contact)
                        {
                            kryptonListBox2.Items.Add(contact.FullName);
                        }
                    }
                }
                else
                {
                    this.Close();
                }
            }

        }



        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            flName = Address_NameLabel.Text;
            if(kryptonCheckBox10.Checked == true)
            {
                loaction = "〒" + loaction;
            }
            else
            {
                loaction = string.Empty;
            }

            if(kryptonCheckBox11.Checked == false)
            {
                MailAddress_User = string.Empty;
                MailAddress_Domain = string.Empty;
            }

            if (kryptonCheckBox12.Checked == false)
            {
                PhoneNumber1 = string.Empty;
                PhoneNumber2 = string.Empty;
                PhoneNumber3 = string.Empty;
                
            }

            if (kryptonCheckBox13.Checked == false)
            {
                FaxNumber1 = string.Empty;
                FaxNumber2 = string.Empty;
                FacNumber3 = string.Empty;

            }
        }

        public int HumanNameOrCompanyName {  get; set; }
        public string flName { get; set; }
        public string loaction { get; set; }

        public string ContactEmailAddress { get; set; }
        public string MailAddress_User { get; set; }
        public string MailAddress_Domain { get; set; }

        public string PhoneNumber1 { get; set; }
        public string PhoneNumber2 { get; set; }
        public string PhoneNumber3 { get; set; }

        public string FaxNumber1 { get; set; }
        public string FaxNumber2 { get; set; }
        public string FacNumber3 { get; set; }

        private void kryptonListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            kryptonButton2.Enabled = true;
            // Outlookアプリケーションのインスタンスを取得
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.NameSpace outlookNs = outlookApp.GetNamespace("MAPI");

            // 連絡先フォルダを取得
            Microsoft.Office.Interop.Outlook.MAPIFolder contactsFolder = outlookNs.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);

            // 検索したい名前
            string targetName = kryptonListBox2.SelectedItem.ToString();

            // 連絡先を検索
            Microsoft.Office.Interop.Outlook.Items contactItems = contactsFolder.Items;
            Microsoft.Office.Interop.Outlook.ContactItem contact = contactItems.Find($"[FullName] = '{targetName}'") as Microsoft.Office.Interop.Outlook.ContactItem;

            if (contact != null)
            {

                loaction = contact.BusinessAddress;

                Address_NameLabel.Text = kryptonListBox2.SelectedItem.ToString();
                kryptonCheckBox10.Text = "所在地:〒" + contact.BusinessAddress;
                kryptonCheckBox11.Text = "メールアドレス:" + contact.Email1Address;
                ContactEmailAddress = contact.Email1Address;
                kryptonCheckBox12.Text = "会社電話番号:" + contact.BusinessTelephoneNumber;
                kryptonCheckBox13.Text = "会社Fax番号:" + contact.BusinessFaxNumber;

                //メールアドレス
                // 変更箇所: kryptonListBox3_SelectedIndexChanged 内の contact.Email1Address を分割してユーザー名とドメイン名を文字列に格納する処理
                // 以下を該当メソッド内の該当行（kryptonRadioButton4.Text = ... の代わり）に置き換えてください。

                // フルアドレスをプロパティに保管（既存プロパティを活用）
                ContactEmailAddress = contact.Email1Address ?? string.Empty;

                // ユーザー名とドメイン名を分割して格納
                MailAddress_User = string.Empty;
                MailAddress_Domain = string.Empty;
                string email = ContactEmailAddress.Trim();

                if (!string.IsNullOrEmpty(email))
                {
                    int at = email.IndexOf('@');
                    if (at > 0 && at < email.Length - 1)
                    {
                        MailAddress_User = email.Substring(0, at);
                        MailAddress_Domain = email.Substring(at + 1);
                    }
                    else
                    {
                        // @ が無い・不正な形式の場合は全体をユーザー部として扱う
                        MailAddress_User = email;
                        MailAddress_Domain = string.Empty;
                    }
                }

                //電話番号
                // 会社電話番号を最大3つのパートに分割して格納
                PhoneNumber1 = PhoneNumber2 = PhoneNumber3 = string.Empty;
                string tel = (contact.BusinessTelephoneNumber ?? string.Empty).Trim();

                if (!string.IsNullOrEmpty(tel))
                {
                    // 数字以外で分割（ハイフンやスペース、括弧などを区切りにする）
                    string[] rawParts = System.Text.RegularExpressions.Regex.Split(tel, @"\D+");
                    System.Collections.Generic.List<string> parts = new System.Collections.Generic.List<string>();
                    foreach (var p in rawParts)
                    {
                        if (!string.IsNullOrEmpty(p)) parts.Add(p);
                    }

                    if (parts.Count >= 3)
                    {
                        PhoneNumber1 = parts[0];
                        PhoneNumber2 = parts[1];
                        PhoneNumber3 = parts[2];
                    }
                    else if (parts.Count == 2)
                    {
                        PhoneNumber1 = parts[0];
                        PhoneNumber2 = parts[1];
                        PhoneNumber3 = string.Empty;
                    }
                    else if (parts.Count == 1)
                    {
                        // 桁数に応じて分割（簡易フォールバック）
                        string digits = parts[0];
                        int len = digits.Length;
                        if (len >= 7)
                        {
                            // 末尾4桁を最後のパートに確保し、先頭を残りで分割
                            int last = 4;
                            int first = Math.Max(2, len - 7); // 最低2桁を先頭に
                            int middle = len - first - last;
                            if (first > 0 && middle >= 0)
                            {
                                PhoneNumber1 = digits.Substring(0, first);
                                if (middle > 0) PhoneNumber2 = digits.Substring(first, middle);
                                PhoneNumber3 = digits.Substring(len - last);
                            }
                            else
                            {
                                PhoneNumber1 = digits;
                            }
                        }
                        else
                        {
                            // 小さい桁数は全体をPhoneNumber1へ
                            PhoneNumber1 = digits;
                        }
                    }

                    //Fax番号
                    ScanFaxNumber();

                }
            }
        }


        public void ScanFaxNumber()
        {
            // Outlookアプリケーションのインスタンスを取得
            Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.NameSpace outlookNs = outlookApp.GetNamespace("MAPI");

            // 連絡先フォルダを取得
            Microsoft.Office.Interop.Outlook.MAPIFolder contactsFolder = outlookNs.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);

            // 検索したい名前
            string targetName = kryptonListBox2.SelectedItem.ToString();

            // 連絡先を検索
            Microsoft.Office.Interop.Outlook.Items contactItems = contactsFolder.Items;
            Microsoft.Office.Interop.Outlook.ContactItem contact = contactItems.Find($"[FullName] = '{targetName}'") as Microsoft.Office.Interop.Outlook.ContactItem;

            // 例: kryptonListBox3_SelectedIndexChanged 内で contact.BusinessFaxNumber を分割して格納
            FaxNumber1 = FaxNumber2 = FacNumber3 = string.Empty;
            string fax = (contact.BusinessFaxNumber ?? string.Empty).Trim();

            if (!string.IsNullOrEmpty(fax))
            {
                // 数字以外で分割（ハイフンやスペース、括弧などを区切りにする）
                string[] rawParts = System.Text.RegularExpressions.Regex.Split(fax, @"\D+");
                System.Collections.Generic.List<string> parts = new System.Collections.Generic.List<string>();
                foreach (var p in rawParts)
                {
                    if (!string.IsNullOrEmpty(p)) parts.Add(p);
                }

                if (parts.Count >= 3)
                {
                    FaxNumber1 = parts[0];
                    FaxNumber2 = parts[1];
                    FacNumber3 = parts[2];
                }
                else if (parts.Count == 2)
                {
                    FaxNumber1 = parts[0];
                    FaxNumber2 = parts[1];
                    FacNumber3 = string.Empty;
                }
                else if (parts.Count == 1)
                {
                    // 桁数に応じて分割（簡易フォールバック）
                    string digits = parts[0];
                    int len = digits.Length;
                    if (len >= 7)
                    {
                        int last = 4;
                        int first = Math.Max(2, len - 7);
                        int middle = len - first - last;
                        if (first > 0 && middle >= 0)
                        {
                            FaxNumber1 = digits.Substring(0, first);
                            if (middle > 0) FaxNumber2 = digits.Substring(first, middle);
                            FacNumber3 = digits.Substring(len - last);
                        }
                        else
                        {
                            FaxNumber1 = digits;
                        }
                    }
                    else
                    {
                        FaxNumber1 = digits;
                    }
                }
            }
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            ContactUpDate();
        }

        private void kryptonRadioButton4_CheckedChanged(object sender, EventArgs e)
        {
            //組織名だった場合
            HumanNameOrCompanyName = 0;
        }

        private void kryptonRadioButton5_CheckedChanged(object sender, EventArgs e)
        {
            //人名だった場合
            HumanNameOrCompanyName = 1;
        }
    }
}
