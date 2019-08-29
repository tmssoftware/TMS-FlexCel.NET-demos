using System;

namespace MainDemo
{
    static class Builder
    {
        internal static void Build(string Proj)
        {
#if(FRAMEWORK40)
            var prj = new Microsoft.Build.Evaluation.Project(Proj);
            try
            {
#else
            Microsoft.Build.BuildEngine.Project prj = new Microsoft.Build.BuildEngine.Project();
            prj.Load(Proj);
#endif

                prj.SetProperty("DefaultTargets", "Build");
                prj.SetProperty("Configuration", "Debug");
                if (!prj.Build()) throw new Exception("Error building project: " + Proj + "\n");
#if(FRAMEWORK40)
            }
            finally
            {
                Microsoft.Build.Evaluation.ProjectCollection.GlobalProjectCollection.UnloadProject(prj);
            }
#endif
        }
    }
}
