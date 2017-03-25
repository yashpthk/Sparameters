using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Numerics;
using System.IO;

namespace CapacitanceExtractorApp
{
    public class DeEmbedHelper
    {
        public void startDeEmbed(string leftPath, string rightPath, string DUTPath, ref string statusMessage)
        {
#if DEBUG
            leftPath = Environment.CurrentDirectory + @"\Data\PORTLINE.S2P";
            rightPath = Environment.CurrentDirectory + @"\Data\PORTLINE.S2P";
            DUTPath = Environment.CurrentDirectory + @"\Data\ONEPF.S2P";
#endif
            string[] leftFile = System.IO.File.ReadAllLines(leftPath);
            string[] rightFile = System.IO.File.ReadAllLines(rightPath);
            string[] DUTFile = System.IO.File.ReadAllLines(DUTPath);

            string[] deembededResult = new string[leftFile.Length];
            deembededResult[0] = "!ILH, Universität Stuttgart, De-embed utility";
            deembededResult[1] = "!Date: " + DateTime.Now.ToString("ddd MMM dd HH:mm:ss yyyy");
            deembededResult[2] = DUTFile[2];
            deembededResult[3] = DUTFile[3];
            deembededResult[4] = DUTFile[4];

            leftFile = leftFile.Skip(5).ToArray();
            rightFile = rightFile.Skip(5).ToArray();
            DUTFile = DUTFile.Skip(5).ToArray();

            if (leftFile.Length == DUTFile.Length && rightFile.Length == DUTFile.Length)
            {
                for (int i = 0; i < leftFile.Length; i++)
                {
                    string freq = leftFile[i].Split('\t')[0];
                    //Get Left fixture S parameters and convert them to Inverse T parameters
                    Complex[,] leftTParams = getParameterMatrix(leftFile[i]);
                    leftTParams = stot(leftTParams);
                    leftTParams = inverse(leftTParams);

                    //Get Right fixture S parameters and convert them to Inverse T parameters
                    Complex[,] rightTParams = getParameterMatrix(rightFile[i]);
                    rightTParams = stot(rightTParams);
                    rightTParams = inverse(rightTParams);
                    
                    //Get DUT S parameters and convert them into T parameters
                    Complex[,] DUTParams = getParameterMatrix(DUTFile[i]);
                    DUTParams = stot(DUTParams);

                    //Deembed
                    DUTParams = Multiply(leftTParams, DUTParams);
                    DUTParams = Multiply(DUTParams, rightTParams);

                    //Convert T to S
                    DUTParams = ttos(DUTParams);

                    deembededResult[i + 5] = freq + " " +
                        DUTParams[0, 0].Magnitude + " " + DUTParams[0, 0].Imaginary + " " +
                        DUTParams[0, 1].Magnitude + " " + DUTParams[0, 1].Imaginary + " " +
                        DUTParams[1, 0].Magnitude + " " + DUTParams[1, 0].Imaginary + " " +
                        DUTParams[1, 1].Magnitude + " " + DUTParams[1, 1].Imaginary + " ";
                }

                using (StreamWriter sw = File.CreateText(Environment.CurrentDirectory + "result.s2p"))
                {
                    foreach (string line in deembededResult)
                    {
                        sw.WriteLine(line);
                    }
                }
            }
            else
            {
                statusMessage = "Every S2P file must have same number of points.";
            }
        }

        private Complex[,] getParameterMatrix(string line)
        {
            var cols = line.Split('\t');
            Complex S11 = new Complex(Convert.ToDouble(cols[1]), Convert.ToDouble(cols[2]));
            Complex S21 = new Complex(Convert.ToDouble(cols[3]), Convert.ToDouble(cols[4]));
            Complex S12 = new Complex(Convert.ToDouble(cols[5]), Convert.ToDouble(cols[6]));
            Complex S22 = new Complex(Convert.ToDouble(cols[7]), Convert.ToDouble(cols[8]));
            Complex[,] leftSParams = new Complex[,] { { S11, S12 }, { S21, S22 } };
            return leftSParams;
        }

        public Complex[,] stot(Complex[,] Sparams)
        {
            Complex det = Determinant(Sparams);
            Complex[,] Tparams = new Complex[2,2];
            Tparams[0, 0] = -(det / Sparams[1, 0]);
            Tparams[0, 1] = Sparams[0, 0] / Sparams[1, 0];
            Tparams[1, 0] = -(Sparams[1, 1] / Sparams[1, 0]);
            Tparams[1, 1] = 1 / Sparams[1, 0];
            return Tparams;
        }

        public Complex[,] ttos(Complex[,] Tparams)
        {
            Complex det = Determinant(Tparams);
            Complex[,] Sparams = new Complex[2, 2];
            Sparams[0, 0] = Tparams[0, 1] / Tparams[1, 1];
            Sparams[0, 1] = det / Tparams[1, 1];
            Sparams[1, 0] = 1 / Tparams[1, 1]; 
            Sparams[1, 1] = -(Tparams[1, 0] / Tparams[1, 1]);
            return Sparams;
        }

        private Complex[,] inverse(Complex[,] Tparams)
        {
            Complex det = Determinant(Tparams);
            Complex[,] inverse = new Complex[2, 2];
            inverse[0, 0] = (Tparams[1, 1] / det);
            inverse[0, 1] = -(Tparams[0, 1] / det);
            inverse[1, 0] = -(Tparams[1, 0] / det);
            inverse[1, 1] = (Tparams[0, 0] / det);
            return inverse;
        }

        private Complex Determinant(Complex[,] param)
        {
            Complex det = new Complex();
            det = (param[0, 0] * param[1, 1]) - (param[0, 1] * param[1, 0]);
            return det;
        }

        private Complex[,] Multiply(Complex[,] A, Complex[,] B)
        {
            Complex[,] R = new Complex[2, 2];
            R[0, 0] = (A[0, 0] * B[0, 0]) + (A[0, 1] * B[1, 0]); //T11*S11 + T12*S21
            R[0, 1] = (A[0, 0] * B[0, 1]) + (A[0, 1] * B[1, 1]); //T11*S12 + T12*S22
            R[1, 0] = (A[1, 0] * B[0, 0]) + (A[1, 1] * B[1, 0]); //T21*S11 + T22*S21
            R[1, 1] = (A[1, 0] * B[0, 1]) + (A[1, 0] * B[1, 1]); //T21*S12 + T22*S22
            return R;
        }
    }
}
