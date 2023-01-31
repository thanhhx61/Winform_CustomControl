using System;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace AITMathRoundControl
{
    public class AITMathRoundTextbox : TextBox
    {
        bool numbersOnly;
        [Browsable(true)]
        [Category("Custom Behavior")]
        [Description("Allows only numbers in the Textbox. If it's set to TRUE, Multiline property must be FALSE")]
        [DisplayName("NumbersOnly")]
        public bool NumbersOnly
        {
            get { return this.numbersOnly; }
            set
            {
                this.numbersOnly = value;
                if (value)
                {
                    this.KeyPress += new KeyPressEventHandler(AITMathRound_KeyPress);
                    base.Multiline = false;
                }
                else
                {
                    this.KeyPress -= new KeyPressEventHandler(AITMathRound_KeyPress);
                }
            }
        }

        // set default NumbersOnly to True
        public AITMathRoundTextbox()
        {
            NumbersOnly = true;

        }

        /// <summary>
        /// Custom Key press event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void AITMathRound_KeyPress(object sender, KeyPressEventArgs e)
        {
            // set cursor to end of number string
            if (e.KeyChar == (char)Keys.Enter)
            {
                Text = MathRound(Text);
                return;
            }
            // accept only number
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.') && (e.KeyChar != '-') && (e.KeyChar != '。') && (e.KeyChar != 'ー') && (e.KeyChar != '．'))
            {
                e.Handled = true;
            }
            // only allow one decimal point, negative sign(fullsize or halfsize)
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1) || (e.KeyChar == '-') && ((sender as TextBox).Text.IndexOf('-') > -1) || (e.KeyChar == 'ー') && ((sender as TextBox).Text.IndexOf('ー') > -1))
            {
                e.Handled = true;
            }

        }

        /// <summary>
        /// Override Multiline
        /// </summary>
        public override bool Multiline
        {
            get { return base.Multiline; }
            set
            {
                if (!this.numbersOnly) base.Multiline = value;
            }
        }

        /// <summary>
        /// Set text 
        /// </summary>
        public override string Text
        {
            get { return base.Text; }
            set
            {
                if (!this.numbersOnly) base.Text = value;
                else
                {
                    Decimal temp = 0;
                    if (!Decimal.TryParse(value, out temp)) base.Text = "";
                    else base.Text = value;
                }
            }
        }

        /// <summary>
        /// Round process
        /// </summary>
        /// <param name="numbox"></param>
        /// <returns></returns>
        private string MathRound(string numbox)
        {
            string NumberMath = string.Empty;
            if (!string.IsNullOrEmpty(numbox) || numbox.Length != 0)
            {
                NumberMath = Math.Round(ConvertStringToDecimal(numbox), MidpointRounding.AwayFromZero).ToString();
            }
            return NumberMath;
        }

        /// <summary>
        /// Convert string to double
        /// </summary>       
        private decimal ConvertStringToDecimal(string valueString)
        {
            // convert Fullsize to Halfsize
            valueString = valueString.Normalize(NormalizationForm.FormKC);
            decimal valueDoube;
            try
            {
                valueDoube = Convert.ToDecimal(valueString);
                return valueDoube;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        /// <summary>
        /// Event key Tab when click button move focus text box
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="keyData"></param>
        /// <returns></returns>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Tab || keyData == Keys.Enter)
            {
                Text = MathRound(Text);           
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        /// <summary>
        /// OnLeave
        /// </summary>
        /// <param name="e"></param>
        protected override void OnLeave(EventArgs e)
        {
            Text = MathRound(Text);
            base.OnLeave(e);
        }
    }
}