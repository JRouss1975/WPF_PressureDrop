using System;
using System.Runtime.Serialization;
using System.ComponentModel;
using MathNet.Numerics;
using BasicInterpolation;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows.Media;

namespace WPF_PressureDrop
{
    //[TypeConverter(typeof(EnumDescriptionTypeConverter))]
    public enum ItemType
    {
        //[Description("Pipe")]
        Pipe,
        Reducer,
        Expander,
        Bend,
        Mitre,
        Tee_Converging,
        Tee_Diverging,
        Butterfly,
        Check,
        Stop,
        Ball,
        Gate,
        Swing,
        Globe,
        Lift,
        Entrance,
        Exit,
        Flange,
        FlangeBlind,
        Component

    }

    [Serializable]
    [DataContract]
    public class HeadLossCalc : Observable, IComparable<HeadLossCalc>, ICloneable
    {

        #region CONSTRUCTORS
        public HeadLossCalc()
        {
            //PropertyChanged += OnPropertyChanged;
        }

        public HeadLossCalc(int id)
        {
            this.ElementId = id;
        }
        #endregion

        #region DATA-TABLES
        /// <summary>
        /// gravitational acceleration [m/sec2].
        /// </summary>
        const double g = 9.80665;

        //range of temperatures for density calculation[F]
        private double[] temps = new double[] { 0.0, 10.0, 20.0, 25, 30.0, 40.0, 50.0, 60.0, 70.0, 80.0, 90.0, 100.0, 110.0, 120.0 };

        //sea water densities for salinity 35 [kg/m3] at pressure 1atm.
        private double[] waterDensities = new double[] { 1028.0, 1027.0, 1024.9, 1023.6, 1022.0, 1018.3, 1014.0, 1009.0, 1003.5, 997.5, 991.0, 984.2, 976.2, 969.3 };

        //sea water viscosities for salinity 35 [cP] at pressure 1atm.
        private double[] waterViscosities = new double[] { 1.906, 1.397, 1.077, 0.959, 0.861, 0.707, 0.594, 0.508, 0.441, 0.393, 0.349, 0.313, 0.280, 0.258 };

        //Bend ratios
        private double[] bendRatio = new double[] { 1.0, 1.5, 2.0, 3.0, 4.0, 6.0, 8.0, 10.0, 12.0, 14.0, 16.0, 20.0 };

        //Bend K values
        private double[] bendK = new double[] { 20.0, 14.0, 12.0, 12.0, 14.0, 17.0, 24.0, 30.0, 34.0, 38.0, 42.0, 50.0 };

        //Mitre Bend angles
        private double[] bendAngle = new double[] { 0.0, 15.0, 30.0, 45.0, 60.0, 75.0, 90.0 };

        //Bend K values
        private double[] mitreBendK = new double[] { 2.0, 4.0, 8.0, 15.0, 25.0, 40.0, 60.0 };
        #endregion

        #region PROPERTIES
        ///<summary>
        /// Line Number
        /// </summary>
        public string Line { get; set; } = "";

        /// <summary>
        /// Notes
        /// </summary>
        private string _Notes = "";
        public string Notes
        {
            get { return _Notes; }
            set
            {
                if (value != _Notes)
                {
                    _Notes = value;
                    NotifyChange("");
                }
            }
        }

        /// <summary>
        /// Element Id
        /// </summary>
        private int _elementId;
        public int ElementId
        {
            get { return _elementId; }
            set
            {
                if (value != _elementId)
                {
                    _elementId = value;
                    NotifyChange("");
                }
            }
        }
        /// <summary>
        /// Start Node.
        /// </summary>
        public Node Node1 { get; set; } = new Node();

        /// <summary>
        /// End Node.
        /// </summary>
        public Node Node2 { get; set; } = new Node();

        /// <summary>
        /// Icon 
        /// </summary>
        public string Icon
        {
            get
            {
                if (this.ElementType == ItemType.Pipe) return "/Resources/Pipe.png";
                if (this.ElementType == ItemType.Bend) return "/Resources/Bend.png";
                if (this.ElementType == ItemType.Expander) return "/Resources/Expander.png";
                if (this.ElementType == ItemType.Reducer) return "/Resources/Reducer.png";
                if (this.ElementType == ItemType.Tee_Converging) return "/Resources/Tee_Converging.png";
                if (this.ElementType == ItemType.Tee_Diverging) return "/Resources/Tee_Diverging.png";
                if (this.ElementType == ItemType.Butterfly) return "/Resources/ButterflyValve.png";
                if (this.ElementType == ItemType.Check) return "/Resources/CheckValve.png";
                if (this.ElementType == ItemType.Flange) return "/Resources/Flange.png";
                if (this.ElementType == ItemType.FlangeBlind) return "/Resources/FlangeBlind.png";
                return "/Resources/MainIcon.png";
            }
        }

        /// <summary>
        /// Temperature degC.
        /// </summary>
        private double _t;
        public double t
        {
            get { return _t; }
            set
            {
                if (value != _t)
                {
                    _t = value;
                    NotifyChange("");
                }
            }
        }

        /// <summary>
        /// ε, absolute roughness or effective height of pipe wall irregularities (mm).
        /// </summary>
        /// 
        private double _epsilon;
        public double epsilon
        {
            get { return _epsilon; }
            set
            {
                if (value != _epsilon)
                {
                    _epsilon = value;
                    NotifyChange("");
                }
            }
        }

        /// <summary>
        /// qh, rate of flow at flowing conditions m3/h,
        /// </summary>
        private double _qh;
        public double qh
        {
            get { return _qh; }
            set
            {
                if (value != _qh)
                {
                    _qh = value;
                    NotifyChange("");
                }
            }
        }

        private double _Qcomb;
        public double Qcomb
        {
            get { return _Qcomb; }
            set
            {
                if (value != _Qcomb)
                {
                    _Qcomb = value;
                    NotifyChange("");
                }
            }
        }

        private double _Qbranch;
        public double Qbranch
        {
            get { return _Qbranch; }
            set
            {
                if (value != _Qbranch)
                {
                    _Qbranch = value;
                    NotifyChange("");
                }
            }
        }

        /// <summary>
        /// Type of fitting
        /// </summary>
        private ItemType _ElementType;
        public ItemType ElementType
        {
            get { return _ElementType; }
            set
            {
                if (value != _ElementType)
                {
                    _ElementType = value;

                    if ((_ElementType == ItemType.Tee_Converging || _ElementType == ItemType.Tee_Diverging) && a == 0)
                        a = 90;

                    NotifyChange("");
                }
            }
        }

        /// <summary>
        /// Length of pipe (m).
        /// </summary>
        private double _L;
        public double L
        {
            get { return _L; }
            set
            {
                if (value != _L)
                {
                    _L = value;
                    NotifyChange("");
                }
            }
        }

        /// <summary>
        /// Internal diameter (m).
        /// </summary>
        private double _d;
        public double d
        {
            get { return _d; }
            set
            {
                if (value != _d)
                {
                    _d = value;
                    NotifyChange("");
                }
            }
        }

        /// <summary>
        /// Reducer/Expander small diamerer or branch diameter d1 in [mm].
        /// </summary>
        public double d1 { get; set; }

        /// <summary>
        /// Reducer/Expander large diamerer or comp diameter d2 in [mm].
        /// </summary>
        public double d2 { get; set; }

        /// <summary>
        /// Bend radious r in [mm].
        /// </summary>
        public double r { get; set; }

        /// <summary>
        /// Number of bends / valves n.
        /// </summary>
        private double _n = 1;
        public double n
        {
            get { return _n; }
            set
            {
                if (value != _n)
                {
                    _n = value;
                    NotifyChange("");
                }
            }
        }

        /// <summary>
        /// Mitre Bend angle a in [deg].
        /// </summary>
        private double _a;
        public double a
        {
            get { return _a; }
            set
            {
                if (value != _a)
                {
                    _a = value;
                    NotifyChange("");
                }
            }
        }

        /// <summary>
        /// Weight of item W in [kg].
        /// </summary>
        private double _W = 0;
        public double W
        {
            get
            {
                if (n >= 1)
                    return _W * n;
                return _W;
            }
            set
            {
                if (value != _W * n)
                {
                    _W = value;
                    NotifyChange("W");
                }
            }
        }
        #endregion

        #region CALCULATED PROPERTIES
        /// <summary>
        /// ρ, water density kg/m3.
        /// </summary>
        public double rho
        {
            get
            {
                double[] p = Fit.Polynomial(temps, waterDensities, 13);
                return Polynomial.Evaluate(t, p);
            }
        }
        /// μ, Dynamic viscocity in cP.
        /// </summary>
        public double me
        {
            get
            {
                double[] p = Fit.Polynomial(temps, waterViscosities, 13);
                return Polynomial.Evaluate(t, p);
            }
        }

        /// <summary>
        /// ν, Kinematic viscocity in cSt.
        /// </summary>
        public double ne
        {
            get { return me / (rho / 1000); }
        }

        /// <summary>
        /// Internal diameter (mm).
        /// </summary>
        private double D
        {
            get
            {
                return d / 1000;
            }
        }

        /// <summary>
        /// Reducer/Expander small diamerer or branch diameter D1 in [m].
        /// </summary>
        private double D1
        {
            get { return d1 / 1000; }
        }

        /// <summary>
        /// Reducer/Expander large or comp diameter d2 diamerer D2 in [m].
        /// </summary>
        private double D2
        {
            get { return d2 / 1000; }
        }

        /// <summary>
        /// Diameter ration for Tee calculation Equation 2-34.
        /// </summary>
        private double Bbranch
        {
            get
            {
                if (D2 > 0)
                    return D1 / D2;
                return 0;
            }
        }

        /// <summary>
        /// q, rate of flow at flowing conditions m3/sec,
        /// </summary>
        private double q
        {
            get { return qh / 3600; }
        }

        /// <summary>
        /// Q, rate of flow at flowing conditions lts/min.
        /// </summary>
        public double Q
        {
            get { return q * 60000; }
        }

        /// <summary>
        /// A, pipe sectional area m2,
        /// </summary>
        public double A
        {
            get
            {
                return (Math.PI / 4) * (D * D);
            }

        }

        /// <summary>
        /// Volume of item V in [liters].
        /// </summary>
        public double Vm
        {
            get
            {
                if (ElementType == ItemType.Reducer || ElementType == ItemType.Expander)
                {
                    double A1 = (Math.PI / 4) * Math.Pow(D1, 2);
                    double A2 = (Math.PI / 4) * Math.Pow(D2, 2);
                    return ((A1 + A2) / 2) * L * 1000 * n;
                }
                return A * L * 1000 * n;
            }
        }

        /// <summary>
        /// v, mean flow velocity m/sec.
        /// </summary>
        public double v
        {
            get
            {
                if (this.ElementType == ItemType.Tee_Converging || this.ElementType == ItemType.Tee_Diverging)
                {
                    return (Math.PI / 4) * ((d2 / 1000) * (d2 / 1000)) > 0 ? (Qcomb / 3600) / ((Math.PI / 4) * ((d2 / 1000) * (d2 / 1000))) : -1;
                }

                return A > 0 ? q / A : -1;
            }
        }

        /// <summary>
        /// Reynolds number (unitless).
        /// </summary>
        public double Re
        {
            get { return (d * v * rho) / me; }
        }

        /// <summary>
        /// Completley turbulent friction factor (Equation 2-8).
        /// </summary>
        public double fT
        {
            get
            {
                return 0.25 / Math.Pow(Math.Log10(epsilon / (3.7 * d)), 2);
            }
        }

        /// <summary>
        /// Serghide’s Solution
        /// </summary>
        private double fs
        {
            get
            {
                double rel_ε = epsilon / d;
                double A = -2 * Math.Log10((rel_ε / 3.7) + (12 / Re));
                double B = -2 * Math.Log10((rel_ε / 3.7) + (2.51 * A / Re));
                double C = -2 * Math.Log10((rel_ε / 3.7) + (2.51 * B / Re));
                return Math.Pow((A - (Math.Pow((B - A), 2) / (C - 2 * B + A))), -2);
            }
        }

        /// <summary>
        /// Colebrook equation friction factor.
        /// </summary>
        public double f
        {
            get
            {
                if (Re < 2000)
                    return 64 / Re;

                if (Re > 4000)
                {
                    double rel_ε = epsilon / d;
                    double f1 = .0000000001;
                    double f2 = 0;
                    do
                    {
                        f2 = f1;
                        f1 = 1 / Math.Pow(-2 * Math.Log10(rel_ε / 3.7 + 2.51 / (Re * Math.Sqrt(f2))), 2);
                    }
                    while (Math.Abs(f1 - f2) > .0000000001);
                    return f1;
                }

                return -1;
            }
        }

        /// <summary>
        /// Loss of static pressure head due to fluid flow [m].
        /// </summary>
        public double hL
        {
            get
            {
                return K * ((v * v) / (2 * g));
            }
        }

        /// <summary>
        /// Loss of static pressure head due to fluid flow [mm].
        /// </summary>
        public double hLm
        {
            get { return hL * 1000; }
        }

        /// <summary>
        /// Resistance coefficient.
        /// </summary>
        public double K
        {
            get
            {
                double k = 0, k1, k2, kb;
                double theta = Math.Abs(2 * Math.Atan((d2 - d1) / (2 * (L * 1000))) * (180 / Math.PI));
                double beta = d1 / d2;
                double rd;
                double r = this.r;

                switch (ElementType)
                {
                    case ItemType.Pipe:
                        k = f * (L / D);
                        break;

                    case ItemType.Reducer:
                        if (theta <= 45)
                        {
                            k1 = (0.8 * Math.Sin((theta / 2) * (Math.PI / 180)) * (1 - Math.Pow(beta, 2)));
                            k2 = k1 / Math.Pow(beta, 4);
                            k = k2;
                        }
                        if (45 < theta && theta <= 180)
                        {
                            k1 = (0.5 * Math.Sqrt(Math.Sin((theta / 2) * (Math.PI / 180))) * (1 - Math.Pow(beta, 2)));
                            k2 = k1 / Math.Pow(beta, 4);
                            k = k2;
                        }
                        break;

                    case ItemType.Expander:
                        if (theta <= 45)
                        {
                            k1 = (2.6 * Math.Sin((theta / 2) * (Math.PI / 180)) * Math.Pow((1 - Math.Pow(beta, 2)), 2));
                            k2 = k1 / Math.Pow(beta, 4);
                            k = k1;
                        }
                        if (45 < theta && theta <= 180)
                        {
                            k1 = Math.Pow((1 - Math.Pow(beta, 2)), 2);
                            k2 = k1 / Math.Pow(beta, 4);
                            k = k1;
                        }
                        break;

                    case ItemType.Bend:
                        if (r / d > 1.25 && r / d < 1.75)
                            rd = 1.5;
                        else
                            rd = Math.Round((r / d), MidpointRounding.ToEven);

                        if (rd >= 1 && rd <= 20)
                        {
                            LinearInterpolation LI0 = new LinearInterpolation(bendRatio, bendK);
                            kb = (double)LI0.Interpolate(rd);
                            if (n == 0)
                                k = kb * fT;
                            else
                                k = ((n - 1) * ((0.25 * Math.PI * fT * (r / d)) + (0.5 * kb)) + kb) * fT;
                        }
                        else
                        {
                            k = -1;
                        }
                        break;

                    case ItemType.Tee_Converging:
                        k = GetTeeKConverging();
                        break;

                    case ItemType.Tee_Diverging:
                        k = GetTeeKDiverging();
                        break;

                    case ItemType.Butterfly:
                        if (d >= 0 && d < 250) k = 45 * fT;
                        if (d >= 250 && d < 400) k = 35 * fT;
                        if (d >= 400) k = 25 * fT;
                        break;

                    case ItemType.Check:
                        if (a == 15)
                        {
                            if (d >= 0 && d < 250) k = 120 * fT;
                            if (d >= 250 && d < 400) k = 90 * fT;
                            if (d >= 400) k = 60 * fT;
                        }
                        else
                        {
                            if (d >= 0 && d < 250) k = 40 * fT;
                            if (d >= 250 && d < 400) k = 30 * fT;
                            if (d >= 400) k = 20 * fT;
                        }
                        break;

                    case ItemType.Lift:
                        k = 600 * fT;
                        break;

                    case ItemType.Globe:
                        k = 340 * fT;
                        break;

                    case ItemType.Gate:
                        k = 8 * fT;
                        break;

                    case ItemType.Swing:
                        k = 100 * fT;
                        break;

                    case ItemType.Ball:
                        k = 3 * fT;
                        break;

                    case ItemType.Entrance:
                        k = 0.78;
                        break;

                    case ItemType.Exit:
                        k = 1;
                        break;

                    case ItemType.Stop:
                        k = 400 * fT;
                        break;

                    case ItemType.Mitre:
                        LinearInterpolation LI1 = new LinearInterpolation(bendAngle, mitreBendK);
                        k = (double)LI1.Interpolate(a) * fT;
                        break;

                    case ItemType.Component:
                        break;

                    default:
                        break;
                }

                if (ElementType == ItemType.Bend)
                    return k;

                return k * n;
            }
        }
        #endregion

        #region METHODS

       
        private double GetTeeKConverging()
        {
            double Krun, Kbranch;
            double Qr = Qbranch / Qcomb;
            if (Qr <= 0.001) Qr = 0;
            double B2 = Math.Pow(Bbranch, 2);
            double C, D, E, F, C1 = 0, D1 = 0, E1 = 0, F1 = 0;

            //C Calculation
            if (B2 <= 0.35)
                C = 1;
            else
            {
                if (Qr <= 0.35)
                    C = 0.9 * (1 - Qr);
                else
                {
                    C = 0.55;
                }
            }

            //Table 2-1 Constants for Equation 2-35
            switch (a)
            {
                case 30.0:
                    D = 1;
                    E = 2;
                    F = 1.74;
                    C1 = 1;
                    D1 = 0;
                    E1 = 1;
                    F1 = 1.74;
                    break;

                case 45.0:
                    D = 1;
                    E = 2;
                    F = 1.41;
                    C1 = 1;
                    D1 = 0;
                    E1 = 1;
                    F1 = 1.41;
                    break;

                case 60.0:
                    D = 1;
                    E = 2;
                    F = 1;
                    C1 = 1;
                    D1 = 0;
                    E1 = 1;
                    F1 = 1;
                    break;

                default:
                    D = 1;
                    E = 2;
                    F = 0;
                    break;
            }


            //Equation 2-35
            Kbranch = C * (1 + D * Math.Pow((Qr * (1 / B2)), 2) - E * Math.Pow(1 - Qr, 2) - F * (1 / B2) * (Math.Pow(Qr, 2)));
            Krun = C1 * (1 + D1 * Math.Pow((Qr * (1 / B2)), 2) - E1 * Math.Pow(1 - Qr, 2) - F1 * (1 / B2) * (Math.Pow(Qr, 2)));

            //Equation 2-36
            if (a == 90)
            {
                Krun = 1.55 * Qr - Math.Pow(Qr, 2);
            }
            if (Qr == 0) Kbranch = 0;

            return Krun + Kbranch;
        }

        private double GetTeeKDiverging()
        {
            double Krun, Kbranch;
            double Qr = Qbranch / Qcomb;
            if (Qr <= 0.001) Qr = 0;
            double B2 = Math.Pow(Bbranch, 2);
            double G = 0, H = 0, J = 0, M = 0;

            //H Calculation
            if (a > 0 && a <= 60) H = 1;
            if (a == 90 && Bbranch <= (2 / 3)) H = 1;
            if (a == 90 && (Bbranch >= 1 || Qr * (1 / B2) <= 2)) H = 0.3;

            //J Calculation
            if (a > 0 && a <= 60) J = 2;
            if (a == 90 && Bbranch <= (2 / 3)) J = 2;
            if (a == 90 && (Bbranch >= 1 || Qr * (1 / B2) <= 2)) J = 0;

            //G Calculation
            if (B2 <= 0.35 && Qr <= 0.6) G = 1.1 - 0.7 * Qr;
            if (B2 <= 0.35 && Qr > 0.6) G = 0.85;
            if (B2 > 0.35 && Qr <= 0.4) G = 1.0 - 0.6 * Qr;
            if (B2 > 0.35 && Qr > 0.4) G = 0.6;

            //M Calculation
            if (B2 <= 0.4 && Qr <= 0.5) M = 0.4;
            if (B2 <= 0.4 && Qr > 0.5) M = 0.4;
            if (B2 > 0.4 && Qr <= 0.5) M = 2 * (2 * Qr - 1);
            if (B2 > 0.4 && Qr > 0.5) M = 0.3 * (2 * Qr - 1);

            //Equation 2-37
            Kbranch = G * (1 + H * Math.Pow((Qr * (1 / B2)), 2) - J * (Qr * (1 / B2)) * Math.Cos((Math.PI * a) / 180));
            Krun = M * Math.Pow(Qr, 2);
            if (Qr == 0) Kbranch = 0;

            return Krun + Kbranch;
        }

        public double Distance(double x1, double x2, double y1, double y2, double z1, double z2)
        {
            return Math.Sqrt(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2) + Math.Pow(z2 - z1, 2));
        }

        public int CompareTo(HeadLossCalc other)
        {
            //if (this.Node1.X == other.Node1.X && this.Node1.Y == other.Node1.Y && this.Node1.Z == other.Node1.Z && this.Node2.X == other.Node2.X && this.Node2.Y == other.Node2.Y && this.Node2.Z == other.Node2.Z) return 0;
            //if (this.Node1.X < other.Node1.X && this.Node1.Y < other.Node1.Y && this.Node1.Z < other.Node1.Z && this.Node2.X < other.Node2.X && this.Node2.Y < other.Node2.Y && this.Node2.Z < other.Node2.Z) return -1;
            //return 1;
            return this.ElementId.CompareTo(other.ElementId);
        }

        public object Clone()
        {
            return this.MemberwiseClone();
        }

        #endregion
    }

    public class Node
    {
        public int NodeId;
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }

    }

    public class ItemEqualityComparer : IEqualityComparer<HeadLossCalc>
    {
        public bool Equals(HeadLossCalc x, HeadLossCalc y)
        {
            return x.Node1.X == y.Node1.X && x.Node1.Y == y.Node1.Y && x.Node1.Z == y.Node1.Z &&
                   x.Node2.X == y.Node2.X && x.Node2.Y == y.Node2.Y && x.Node2.Z == y.Node2.Z &&
                   x.L == y.L &&
                   x.d == y.d &&
                   x.ElementType == y.ElementType;
        }

        public int GetHashCode(HeadLossCalc obj)
        {
            return -1;
        }
    }
}