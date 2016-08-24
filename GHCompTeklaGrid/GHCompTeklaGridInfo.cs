using System;
using System.Drawing;
using Grasshopper.Kernel;

namespace GHCompTeklaGrid
{
    public class GHCompTeklaGridInfo : GH_AssemblyInfo
    {
        public override string Name
        {
            get
            {
                return "TeklaGrid";
            }
        }
        public override Bitmap Icon
        {
            get
            {
                return null;
            }
        }
        public override string Description
        {
            get
            {
                return "Exemple of interoperability between Grasshopper and Tekla";
            }
        }
        public override Guid Id
        {
            get
            {
                return new Guid("1f5caa0a-6dca-46d7-8274-96038f50e24a");
            }
        }

        public override string AuthorName
        {
            get
            {
                //Return a string identifying you or your company.
                return "Brice Canteneur for EIFFAGE Métal";
            }
        }
        public override string AuthorContact
        {
            get
            {
                //Return a string representing your preferred contact details.
                return "brice.canteneur@eiffage.com";
            }
        }
    }
}
