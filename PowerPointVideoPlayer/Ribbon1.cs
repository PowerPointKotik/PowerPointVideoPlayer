private void button1_Click(object sender, RibbonControlEventArgs e)
{
    using (var window = new OpenGLWindow())
    {
        window.Run(60.0);
    }
}