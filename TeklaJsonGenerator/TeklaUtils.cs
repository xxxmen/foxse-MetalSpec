using System;
using System.Collections;
using System.Collections.Generic;
using Tekla.Structures.Model;

namespace TeklaJsonGenerator
{
    class TeklaUtils
    {
        internal static string GetConfigFilePath()
        {
            return "X:\\Apps\\Tekla\\applications\\JsonGenerator\\TeklaMetalSpecConfig.txt";
        }

        internal static void GetDescendants(Assembly assy, ref List<Part> outParts)
        {
            ModelObject mainPart = assy.GetMainPart();
            if (mainPart != null)
            {
                outParts.Add(mainPart as Part);

                ModelObjectEnumerator mainPartChildren = mainPart.GetChildren();

                while (mainPartChildren.MoveNext())
                {
                    ModelObject mainPartChild = mainPartChildren.Current as ModelObject;
                    if (mainPartChild != null)
                    {
                        if (mainPartChild is Part)
                        {
                            outParts.Add(mainPartChild as Part);
                        }
                    }
                }
            }

            ArrayList secondaries = assy.GetSecondaries();
            if (secondaries.Count > 0)
            {
                foreach (var secondary in secondaries)
                {
                    Part part = secondary as Part;
                    if (part != null)
                    {
                        outParts.Add(part);

                        ModelObjectEnumerator secondaryChildren = part.GetChildren();

                        while (secondaryChildren.MoveNext())
                        {
                            ModelObject secondaryChild = secondaryChildren.Current as ModelObject;
                            if (secondaryChild != null)
                            {
                                if (secondaryChild is Part)
                                {
                                    outParts.Add(secondaryChild as Part);
                                }
                            }
                        }
                    }
                }
            }

            ArrayList subAssies = assy.GetSubAssemblies();
            if (subAssies.Count > 0)
            {
                foreach (var item in subAssies)
                {
                    Assembly subAssy = item as Assembly;
                    if (subAssy != null)
                    {
                        GetDescendants(subAssy, ref outParts);
                    }
                }
            }
        }
    }
}
