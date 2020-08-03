using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;

namespace TITHES
{
    public partial class Tithes : Form
    {
        public Tithes()
        {
            InitializeComponent();
        }

            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=|DataDirectory|\TithesDB1.mdb");
            decimal tithe, generaloffering, sundayschool, missions, youthministry, childrensministry, radioministry, buildingfund, other, total;
            decimal titheTotal, generalofferingTotal, sundayschoolTotal, missionsTotal, youthministryTotal, childrensministryTotal, radioministryTotal, buildingfundTotal, otherTotal, memberTotal;
            int Acct;
            string check, MemberName, TitheTotalyear, FormattedDate, month, day, year, date;

        private void Tithes_Load(object sender, EventArgs e)
        {
            this.DesktopLocation = new Point(420, 20);
            ReadMemberName();
            
        }


        private void button1_Click(object sender, EventArgs e)
        {
         
            int error = 0;
            labelError.Visible = false;
            if (textName.Text != "")
                ++error;

            if (Int32.TryParse(textAcct.Text.Trim(), out Acct))
            {
                if (Acct >= 0)
                    ++error;
            }
            else
            {
                Acct = 0;
                ++error;
            }
            if (textTithe.Text == "")
            {
                ++error; 
                tithe = 0;
            }
            else
                if (Decimal.TryParse(textTithe.Text.Trim(), out tithe))
                    if (tithe >= 0)
                    ++error; 
            if (textOffering.Text == "")
            {
                ++error; 
                generaloffering = 0;
            }
            else
                if (Decimal.TryParse(textOffering.Text.Trim(), out generaloffering))
                    if (generaloffering >= 0)
                    ++error; 
            if (textSS.Text == "")
            {
                ++error; 
                sundayschool = 0;
            }
            else
                if (Decimal.TryParse(textSS.Text.Trim(), out sundayschool))
                    if (sundayschool >= 0)
                    ++error; 
            if (textMission.Text == "")
            {
                ++error; 
                missions = 0;
            }
            else
                if (Decimal.TryParse(textMission.Text.Trim(), out missions))
                    if (missions >= 0)
                    ++error; 
            if (textYouth.Text == "")
            {
                ++error; 
                youthministry = 0;
            }
            else
                if (Decimal.TryParse(textYouth.Text.Trim(), out youthministry))
                    if (youthministry >= 0)
                    ++error; 
            if (textChildrens.Text == "")
            {
                ++error; 
                childrensministry = 0;
            }
            else
                if (Decimal.TryParse(textChildrens.Text.Trim(), out childrensministry))
                    if (childrensministry >= 0)
                    ++error; 
            if (textRadio.Text == "")
            {
                ++error; 
                radioministry = 0;
            }
            else
                if (Decimal.TryParse(textRadio.Text.Trim(), out radioministry))
                    if (radioministry >= 0)
                    ++error; 
            if (textBuilding.Text == "")
            {
                ++error;
                buildingfund = 0;
            }
            else
                if (Decimal.TryParse(textBuilding.Text.Trim(), out buildingfund))
                    if (buildingfund >= 0)
                    ++error;
            if (textOther.Text == "")
            {
                ++error;
                other = 0;
            }
            else
                if (Decimal.TryParse(textOther.Text.Trim(), out other))
                    if (other >= 0)
                    ++error;
            if (error == 11)
            {
                FormattedDate = TextDate.Text;
                string[] dateComponents = FormattedDate.Split('/');
                month = dateComponents[0].Trim(); ;
                day = dateComponents[1].Trim();
                year = dateComponents[2].Trim();
                date = month + day + year;
                TitheTotalyear = "TitheTotal" + year;

                ReadTotals();

                titheTotal += tithe;
                generalofferingTotal += generaloffering;
                sundayschoolTotal += sundayschool;
                missionsTotal += missions;
                youthministryTotal += youthministry;
                childrensministryTotal += childrensministry;
                radioministryTotal += radioministry;
                buildingfundTotal += buildingfund;
                otherTotal += other;
                MemberName = textName.Text;
                total = tithe + generaloffering + sundayschool + missions + youthministry + childrensministry + radioministry + buildingfund + other;
                labelTotal.Text = total.ToString("C");
                check = "false";
                
                int count = textName.Items.Count;
                for (int z = 0; z < count; ++z)
                {
                    string existingmember = textName.Items[z].ToString();
                    if (MemberName == existingmember)
                    {
                        ReadMemberTotal();
                        memberTotal += total;
                        check = "true";
                    }
                }
                
                

                WriteTextFile();

                WriteDatabase();

                ReadMemberName();
            }
            else
                labelError.Visible = true;
                labelError.Text = "Please be Sure to Enter All Information Accurately";
            
        }


        private void WriteTextFile()
        {
            
            string filename = MemberName + date + "Offering.txt";
            FileStream outFile = new FileStream(filename,
                        FileMode.Create, FileAccess.Write);
            StreamWriter writer = new StreamWriter(outFile);
            writer.WriteLine("Name: " + MemberName);
            writer.WriteLine("Date: " + FormattedDate);
            writer.WriteLine("Acct#: " + textAcct.Text);
            writer.WriteLine("Tithe: " + tithe);
            writer.WriteLine("Offering: " + generaloffering);
            writer.WriteLine("Sunday school: " + sundayschool);
            writer.WriteLine("Missions: " + missions);
            writer.WriteLine("Youth Ministry: " + youthministry);
            writer.WriteLine("Childrens Ministry: " + childrensministry);
            writer.WriteLine("Radio Ministry: " + radioministry);
            writer.WriteLine("Building Fund: " + buildingfund);
            writer.WriteLine("Other: " + other);
            writer.WriteLine("Total: " + total);

            writer.Close();
            outFile.Close();


 
        }

        private void WriteDatabase()
        {
                 
            
            OleDbCommand cmmd = new OleDbCommand();
            OleDbCommand cmmd2 = new OleDbCommand();
            OleDbCommand cmmd6 = new OleDbCommand();
            cmmd.Connection = conn;
            cmmd2.Connection = conn;
            cmmd6.Connection = conn;             
            
            cmmd.CommandText = "INSERT INTO TitheTotals ([MemberName], [Acct#], [" + TitheTotalyear + "]) VALUES('" + MemberName + "', '" + Acct + "', '" + total + "')";
            cmmd2.CommandText = "Update Totals Set Tithe=?, GeneralOffering=?, SundaySchool=?, Missions=?, YouthMinistry=?, [Children'sMinistry]=?, RadioTVMinistry=?, BuildingFund=?, Other=? WHERE YearOfTithe=?";
            cmmd2.Parameters.AddWithValue("@Tithe",titheTotal);
            cmmd2.Parameters.AddWithValue("@GeneralOffering", generalofferingTotal);
            cmmd2.Parameters.AddWithValue("@SundaySchool", sundayschoolTotal);
            cmmd2.Parameters.AddWithValue("@Missions", missionsTotal);
            cmmd2.Parameters.AddWithValue("@YouthMinistry", youthministryTotal);
            cmmd2.Parameters.AddWithValue("@[Children'sMinistry]", childrensministryTotal);
            cmmd2.Parameters.AddWithValue("@RadioTVMinistry", radioministryTotal);
            cmmd2.Parameters.AddWithValue("@BuildingFund", buildingfundTotal);
            cmmd2.Parameters.AddWithValue("@Other", otherTotal);
            cmmd2.Parameters.AddWithValue("@YearOfTithe", year);
            cmmd6.CommandText = "Update TitheTotals Set [" + TitheTotalyear + "]=? Where [MemberName]=?;";
            cmmd6.Parameters.AddWithValue("@["+TitheTotalyear+"]", memberTotal);
            cmmd6.Parameters.AddWithValue("@MemberName", MemberName);
            conn.Open();

            
            if (conn.State == ConnectionState.Open)
            {               
                    if (check == "true")
                    {
                        cmmd6.ExecuteNonQuery();
                    }
                    else
                        cmmd.ExecuteNonQuery();
                        cmmd2.ExecuteNonQuery();
                        MessageBox.Show("Your Tithe Has Been Recorded.","Important Message");                
                    conn.Close();
                
            }
        }

        private void ReadTotals()
        {
            
            
            OleDbCommand cmmd4 = new OleDbCommand();
            cmmd4.Connection = conn;
            cmmd4.CommandText = "Select Tithe, GeneralOffering, SundaySchool, Missions, YouthMinistry, [Children'sMinistry], RadioTVMinistry, BuildingFund, Other From [Totals] Where YearOfTithe="+year+"";
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                
            OleDbDataReader reader = cmmd4.ExecuteReader();
            while (reader.Read())
            {
                titheTotal = Convert.ToDecimal(reader[0]);
                generalofferingTotal = Convert.ToDecimal(reader[1]);
                sundayschoolTotal = Convert.ToDecimal(reader[2]);
                missionsTotal = Convert.ToDecimal(reader[3]);
                youthministryTotal = Convert.ToDecimal(reader[4]);
                childrensministryTotal = Convert.ToDecimal(reader[5]);
                radioministryTotal = Convert.ToDecimal(reader[6]);
                buildingfundTotal = Convert.ToDecimal(reader[7]);
                otherTotal = Convert.ToDecimal(reader[8]);
            }
                    conn.Close();
            }            
        }
        
        private void ReadMemberName()
        {
            textName.Items.Clear();
            conn.Open();
            OleDbCommand cmmd3 = new OleDbCommand();
            cmmd3.Connection = conn;
            cmmd3.CommandText = "Select MemberName From TitheTotals";
            OleDbDataReader reader = cmmd3.ExecuteReader();
            while (reader.Read())
            {
                textName.Items.Add(reader[0].ToString());
            }
            conn.Close();
        }

        private void ReadMemberTotal()
        {
            
            OleDbCommand cmmd5 = new OleDbCommand();
            cmmd5.Connection = conn;
            cmmd5.CommandText = "Select [" + TitheTotalyear + "] From TitheTotals Where MemberName=?";
            cmmd5.Parameters.AddWithValue("@MemberName",MemberName);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {

                OleDbDataReader reader = cmmd5.ExecuteReader();
                while (reader.Read())
                {
                    memberTotal = Convert.ToDecimal(reader[0]);
                }
                conn.Close();
            }
        }

    

        private void btnClear_Click(object sender, EventArgs e)
        {
            textAcct.Text = "";
            TextDate.Text = "";
            textName.Text="";
            textName.Focus();
            textTithe.Text = "";
            textOffering.Text = "";
            textSS.Text = "";
            textMission.Text = "";
            textYouth.Text = "";
            textChildrens.Text = "";
            textRadio.Text = "";
            textBuilding.Text = "";
            textOther.Text = "";
            labelTotal.Text = "TOTAL";
        }

    }
}
