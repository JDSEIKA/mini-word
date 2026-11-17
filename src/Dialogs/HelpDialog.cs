using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordPadClone.Dialogs
{
    internal sealed class HelpDialog : Form
    {
        public HelpDialog()
        {
            Text = "Ayuda de Mini Word";
            Font = new Font("Segoe UI", 9f);
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(560, 420);
            MinimumSize = new Size(420, 300);

            TextBox text = new TextBox();
            text.Multiline = true;
            text.ReadOnly = true;
            text.ScrollBars = ScrollBars.Vertical;
            text.Dock = DockStyle.Fill;
            text.BackColor = Color.White;
            text.Text =
                "Mini Word: guía rápida\r\n\r\n"
                + "ARCHIVO\r\n"
                + "Nuevo, abrir, guardar, guardar como, imprimir y vista previa. "
                + "El autoguardado se configura desde el botón Autoguardado y permite activarlo y elegir cada cuántos minutos se repite.\r\n\r\n"
                + "INICIO\r\n"
                + "Portapapeles, fuente, tamaño, negrita, cursiva, subrayado, color, alineación, viñetas, salto de página y encabezado.\r\n\r\n"
                + "INSERTAR\r\n"
                + "Tablas con fusión y división de celdas; imágenes flotantes con tiradores para mover y redimensionar; fecha; "
                + "cuadros de texto; formas (rectángulo, círculo, triángulo, estrella, flecha, rombo y corazón) y encabezado/pie.\r\n\r\n"
                + "DISPOSICIÓN\r\n"
                + "Configurar página y márgenes, columnas periodísticas y saltos de sección.\r\n\r\n"
                + "REVISAR\r\n"
                + "Ortografía del sistema, panel de comentarios y buscar/reemplazar.\r\n\r\n"
                + "HERRAMIENTAS\r\n"
                + "Ortografía, Autoguardado, esta Ayuda y Acerca de.\r\n\r\n"
                + "OBJETOS FLOTANTES\r\n"
                + "Haga clic en una imagen, cuadro de texto o forma para seleccionarlo. "
                + "Arrastre el objeto para moverlo y use los ocho tiradores para cambiar su tamaño. "
                + "Doble clic sobre un cuadro o forma permite editar su texto. Supr lo elimina.";
            Controls.Add(text);

            Panel bottom = new Panel();
            bottom.Dock = DockStyle.Bottom;
            bottom.Height = 44;
            Button close = new Button();
            close.Text = "Cerrar";
            close.DialogResult = DialogResult.OK;
            close.Location = new Point(464, 9);
            close.Width = 80;
            bottom.Controls.Add(close);
            Controls.Add(bottom);
            AcceptButton = close;
            CancelButton = close;
        }
    }
}
