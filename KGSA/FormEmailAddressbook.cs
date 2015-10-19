using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlServerCe;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;

namespace KGSA
{
    public partial class FormEmailAddressbook : Form
    {
        public FormEmailAddressbook()
        {
            InitializeComponent();
            UpdateEmailsDatagrid();
        }

        DataSet ds;
        SqlCeDataAdapter da;
        SqlCeConnection con = new SqlCeConnection(FormMain.SqlConStr);
        private void UpdateEmailsDatagrid()
        {

            ds = new DataSet();
            con.Open();
            da = new SqlCeDataAdapter("Select * from tblEmail", con);
            var cmdBldr = new SqlCeCommandBuilder(da);
            da.Fill(ds, "tblEmail");
            bindingSource1.DataSource = ds;
            bindingSource1.DataMember = "tblEmail";
            dataGridView1.DataSource = bindingSource1;
            bindingNavigator1.BindingSource = bindingSource1;

            con.Close();
        }

        public bool Add(string nameEmail, string addressEmail, string typeEmail, bool quickEmail)
        {
            nameEmail = nameEmail.Trim();

            if (CheckDuplicate(addressEmail))
                return false; // Denne adresse finnes allerede

            try
            {
                con.Open();
                using (SqlCeCommand cmd = new SqlCeCommand("insert into tblEmail(Name, Address, Type, Quick) values (@Val1, @val2, @val3, @val4)", con))
                {
                    cmd.Parameters.AddWithValue("@Val1", nameEmail);
                    cmd.Parameters.AddWithValue("@Val2", addressEmail);
                    cmd.Parameters.AddWithValue("@Val3", typeEmail);
                    cmd.Parameters.AddWithValue("@Val4", quickEmail);
                    cmd.CommandType = System.Data.CommandType.Text;
                    cmd.ExecuteNonQuery();
                }
                con.Close();
                return true;
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
                return false;
            }
        }

        private bool CheckDuplicate(string address)
        {
            con.Open();
            SqlCeCommand cmd = con.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*) FROM tblEmail WHERE Address = '" + address + "'";
            int result = ((int)cmd.ExecuteScalar());
            con.Close();
            if (result == 0)
                return false;
            return true;
        }

        /// <summary>
        /// Knapp for å legge til en adresse
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string addressEmail = textBoxAdresse.Text;
                string nameEmail = textBoxNavn.Text;
                if (addressEmail.Length < 3 || nameEmail.Length < 1)
                    return;

                string addressValidated = addressEmail.ToLower();
                addressValidated = new MailAddress(addressValidated).ToString();

                Add(nameEmail, addressValidated, "Full", false); // Legg til i databasen

                UpdateEmailsDatagrid();

                textBoxNavn.Text = "";
                textBoxAdresse.Text = "";
                textBoxNavn.Focus();
            }
            catch (FormatException)
            {
                Logg.Log("Epost adresse ugyldig.", Color.Red);
                textBoxAdresse.Focus();
                textBoxAdresse.SelectAll();
            }
            catch(Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void SaveEmailsDatagrid()
        {
            try
            {
                if (dataGridView1.CurrentRow != null)
                    dataGridView1.CurrentRow.DataGridView.EndEdit();
                dataGridView1.EndEdit();
                bindingSource1.EndEdit();

                con.Open();
                da.Update(ds, "tblEmail");
                con.Close();

                UpdateEmailsDatagrid();
                Logg.Log("Epost adresser oppdatert.", Color.Green);
            }
            catch (Exception ex)
            {
                Logg.Unhandled(ex);
            }
        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            SaveEmailsDatagrid();
        }

        private void textBoxAdresse_Enter(object sender, EventArgs e)
        {
            ActiveForm.AcceptButton = button1;
        }

        private void textBoxAdresse_Leave(object sender, EventArgs e)
        {
            ActiveForm.AcceptButton = null;
        }

        private void FormEmail_FormClosing(object sender, FormClosingEventArgs e)
        {
            Logg.Status("E-post adresse vindu avsluttet.");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveEmailsDatagrid();
        }

        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            SaveEmailsDatagrid();
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            Logg.Debug("DataError (ColumnIndex " + e.ColumnIndex + ", RowIndex " + e.RowIndex + "): " + e.Exception.Message + " Context: " + e.Context);
        }
    }
}
