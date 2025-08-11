// TARGET:dummy.exe
// START_IN:
using LoginPI.Engine.ScriptBase;

public class Default : ScriptBase
{
    void Execute() 
    {
        START();
        Wait(2); 
        STOP();
    }
}