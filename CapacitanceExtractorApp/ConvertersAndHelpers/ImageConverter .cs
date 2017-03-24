using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace CapacitanceExtractorApp
{
    public class ImageConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            BitmapImage image = new BitmapImage();
            image.BeginInit();
            if (value != null)
                image.UriSource = new Uri(value.ToString(), UriKind.RelativeOrAbsolute);
            else
                image.UriSource = new Uri("pack://application:,,,/CapacitanceExtractorApp;component/Resources/blank.png", UriKind.RelativeOrAbsolute);
            image.CacheOption = BitmapCacheOption.OnLoad;
            image.EndInit();
            return image;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
