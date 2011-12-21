using System;
using System.Drawing;
using System.Windows.Forms;
using stdole;

public class AxHost2 : AxHost
{
    public AxHost2()
        : base(null)
    {
    }
    public new static IPictureDisp GetIPictureDispFromPicture(Image image)
    {
        return (IPictureDisp)AxHost.GetIPictureDispFromPicture(image);
    }
}