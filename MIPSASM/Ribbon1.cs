using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace MIPSASM
{
    public partial class Ribbon1
    {
        enum CommandTypes
        {
            RType,
            IType,
            JType
        }
        Dictionary<string, string> instr = new Dictionary<string, string>()
        {
            { "LW",   "100011" },
            { "SW",   "101011" },
            { "ADDI", "001000" },
            { "ADD",  "100000" },
            { "SUBI", "001010" },
            { "SUB",  "100010" },
            { "MUL", "011000" },
            { "J",    "000010" },
            { "BNE",  "000101" },
            { "BEQ",  "000100" },
            { "LS",   "000000" },
            { "RS",   "000010" }
        };
        Dictionary<string, string> registry = new Dictionary<string, string>()
        {
            {"$zero", "00000"},
            {"$at",   "00001"},
            {"$v0",   "00010"},
            {"$v1",   "00011"},
            {"$a0",   "00100"},
            {"$a1",   "00101"},
            {"$a2",   "00110"},
            {"$a3",   "00111"},
            {"$t0",   "01000"},
            {"$t1",   "01001"},
            {"$t2",   "01010"},
            {"$t3",   "01011"},
            {"$t4",   "01100"},
            {"$t5",   "01101"},
            {"$t6",   "01110"},
            {"$t7",   "01111"},
            {"$s0",   "10000"},
            {"$s1",   "10001"},
            {"$s2",   "10010"},
            {"$s3",   "10011"},
            {"$s4",   "10100"},
            {"$s5",   "10101"},
            {"$s6",   "10110"},
            {"$s7",   "10111"},
            {"$t8",   "11000"},
            {"$t9",   "11001"},
            {"$k0",   "11010"},
            {"$k1",   "11011"},
            {"$gp",   "11100"},
            {"$sp",   "11101"},
            {"$fp",   "11110"},
            {"$ra",   "11111"},
        };
        Dictionary<string, CommandTypes> types = new Dictionary<string, CommandTypes>()
        {
            { "LW",   CommandTypes.IType },
            { "SW",   CommandTypes.IType },
            { "ADDI", CommandTypes.IType },
            { "ADD",  CommandTypes.RType },
            { "SUBI", CommandTypes.IType },
            { "SUB",  CommandTypes.RType },
            { "MUL", CommandTypes.IType },
            { "J",    CommandTypes.JType },
            { "BNE",  CommandTypes.IType },
            { "BEQ",  CommandTypes.IType },
            { "LS",   CommandTypes.IType },
            { "RS",   CommandTypes.IType }
        };

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            string selectText = string.Empty;
            Selection wordSelection = Globals.ThisAddIn.Application.Selection;
            if (wordSelection != null && wordSelection.Range != null)
            {
                selectText = wordSelection.Text;
            }
            string result = "";
            foreach (string instr in selectText.Split('\r'))
            {
                result += getInstrBinary(instr);
                result += "\r";
            }
            wordSelection.Text = result;
        }

        private string getInstrBinary(string instr)
        {
            if (string.IsNullOrEmpty(instr)) return string.Empty;
            switch (types[instr.Split()[0]])
            {
                case CommandTypes.RType:
                    return getRTypeBinary(instr);
                case CommandTypes.IType:
                    return getITypeBinary(instr);
                case CommandTypes.JType:
                    return getJTypeBinary(instr);
            }
            throw new ArgumentException();
        }

        private string getRTypeBinary(string instr)
        {
            string result = ""; // new string
            instr = instr.Replace(",", " ").Replace("(", " ").Replace(')', ' ').Trim(); // change all chars to spaces
            instr = Regex.Replace(instr, @"\s+", " "); // evil regex to change double spaces to one
            result += "000000"; // opcode 6 bits
            result += "_";
            result += registry[instr.Split()[2]]; // rs 5 bits
            result += "_";
            result += registry[instr.Split()[3]]; // rt 5 bits
            result += "_";
            result += registry[instr.Split()[1]]; // rd 5 bits
            result += "_";
            result += "00000"; // shamt 5 bits
            result += "_";
            result += this.instr[instr.Split()[0]]; // funct 6 bits
            return result;
        }

        private string getITypeBinary(string instr)
        {
            string result = ""; // new string
            instr = instr.Replace(",", " ").Replace("(", " ").Replace(')', ' ').Trim(); // change all chars to spaces
            instr = Regex.Replace(instr, @"\s+", " "); // evil regex to change double spaces to one
            result += this.instr[instr.Split()[0]]; // opcode 6 bits
            result += "_";
            if (instr.Split()[0] == "LW" || instr.Split()[0] == "SW")
            {
                result += registry[instr.Split()[3]]; // rs 5 bits
                result += "_";
                result += registry[instr.Split()[1]]; // rd 5 bits
                result += "_";
                result += get16Bit(Convert.ToInt32(instr.Split()[2].Replace("0x", ""), 16)); // imm 16 bits
            }
            else
            {
                result += registry[instr.Split()[2]]; // rs 5 bits
                result += "_";
                result += registry[instr.Split()[1]]; // rd 5 bits
                result += "_";
                result += get16Bit(Convert.ToInt32(instr.Split()[3])); // imm 16 bits
            }
            return result;
        }

        private string getJTypeBinary(string instr)
        {
            string result = ""; // new string
            result += this.instr[instr.Split()[0]]; // opcode 6 bits
            result += "_";
            result += get26Bit(Convert.ToInt32(instr.Split()[1])); // offset 26 bits
            return result;
        }

        private string get16Bit(int number)
        {
            if (number < 0)
            {
                number = (1 << 16) + number;
            }

            string binary = Convert.ToString(number, 2);

            binary = binary.PadLeft(16, '0');

            return binary;
        }

        private string get26Bit(int number)
        {
            if (number < 0)
            {
                number = (1 << 26) + number;
            }

            string binary = Convert.ToString(number, 2);

            binary = binary.PadLeft(26, '0');

            return binary;
        }
    }
}
