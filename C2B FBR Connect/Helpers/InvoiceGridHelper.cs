using System.Collections.Generic;
using System.Windows.Forms;
using C2B_FBR_Connect.Models;
using System.Drawing;

namespace C2B_FBR_Connect.Helpers
{
  public static class InvoiceGridHelper
  {
    public static void FormatDataGridView(DataGridView dgv, List<Invoice> invoices)
    {
      dgv.DataSource = null;
      dgv.Rows.Clear();
      dgv.Columns.Clear();
      if (invoices == null || invoices.Count == 0)
        return;
      dgv.DataSource = invoices;
      // Color code rows by status
      foreach (DataGridViewRow row in dgv.Rows)
      {
        var status = row.Cells["Status"]?.Value?.ToString();
        switch (status)
        {
          case "Uploaded":
            row.DefaultCellStyle.BackColor = Color.FromArgb(200, 250, 205);
            break;
          case "Failed":
            row.DefaultCellStyle.BackColor = Color.FromArgb(255, 205, 210);
            break;
          case "Pending":
            row.DefaultCellStyle.BackColor = Color.FromArgb(255, 248, 225);
            break;
        }
      }
    }

    public static void HideColumn(DataGridView dgv, string columnName)
    {
      if (dgv.Columns[columnName] != null)
        dgv.Columns[columnName].Visible = false;
    }

    public static void SetColumnHeader(DataGridView dgv, string columnName, string headerText, int width = -1, string format = null)
    {
      if (dgv.Columns[columnName] != null)
      {
        dgv.Columns[columnName].HeaderText = headerText;
        if (width > 0)
          dgv.Columns[columnName].Width = width;
        if (!string.IsNullOrEmpty(format))
          dgv.Columns[columnName].DefaultCellStyle.Format = format;
      }
    }
  }
}
