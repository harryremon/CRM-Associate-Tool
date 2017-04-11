namespace CRMAssociateTool.Console.Helpers
{
    public static class StringHelpers
    {
        public static void Swap(ref string firstString,ref string secondString)
        {
            var temp = firstString;
            firstString = secondString;
            secondString = temp;
        }
    }
}