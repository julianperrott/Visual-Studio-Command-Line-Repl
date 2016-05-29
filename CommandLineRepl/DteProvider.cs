namespace CommandLineRepl
{
    public interface IDteProvider
    {
        EnvDTE80.DTE2 Dte
        {
            get;
        }
    }
}