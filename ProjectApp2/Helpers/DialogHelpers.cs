using Microsoft.Win32;

namespace ProjectApp2.Helpers
{
    public class DialogHelpers
    {
        public static string GetFileName(string filterName)
        {
            string selectedFile = string.Empty;
            OpenFileDialog dlg = new OpenFileDialog
            {
                Multiselect = false,
                ValidateNames = true,
                Filter = filterName
            };
            bool? dialogResult = dlg.ShowDialog();
            if (dialogResult == true)
            {
                selectedFile = dlg.FileName;
            }
            return selectedFile;
        }
    }

}
