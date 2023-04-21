using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Reflection;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace Amend_Connection_Reference
{
    // Do not forget to update version number and author (company attribute) in AssemblyInfo.cs class
    // To generate Base64 string for Images below, you can use https://www.base64-image.de/
    [Export(typeof(IXrmToolBoxPlugin)),
        ExportMetadata("Name", "Amend Connection Reference"),
        ExportMetadata("Description", "This is a tool to update Connection Reference of Cloud Flows in Dataverse."),
        // Please specify the base64 content of a 32x32 pixels image
        ExportMetadata("SmallImageBase64", "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAABGdBTUEAALGPC/xhBQAAACBjSFJNAAB6JgAAgIQAAPoAAACA6AAAdTAAAOpgAAA6mAAAF3CculE8AAAABmJLR0QA/wD/AP+gvaeTAAAACXBIWXMAAAsTAAALEwEAmpwYAAAAB3RJTUUH5wQVCwED2ZUP7QAABG1JREFUWMO1139oXWcdx/HXc25uEpMms7rarW4wZ0ZrRaai7RjbbKTFzQ7WXzAmGCrTyTaLm39I/3FrEcYYmxaKRSrimj+KhTlaSUHcH+20oqUTGZuhq02bORWxtsPkNsm9N/c8/nFumpvk3uSaxi8cOOc85zzvz/P5fp/nOSe4nnj1OHxYjIfF+AUhHJAkz4ixYNvmprpIrktAjMR4pxjvQZsYn5Sm3xNCR1Xc/1kAhHAD2qpXrWJ8WpruEUJnMyJC06Czw1NneawSwmrFYo8LQxtMTu6Y1VdZCPskyV4xXp0vHQsLmAYvx0Zsw3qsUiq1uXiBcrnem02JmF9ABm/DZnwbd6H1WnupZB4BtSL2iHGsnoj6NXB2eAp+M36Aftw3A95c5MX4lDT9pkKBY79qQsC05atxCE+g838EzxbxFV1dy6XpnMaWBvCP42B11EsRHWLM12uol4LleGEJ4YRwSj5/WTIXN30nG33A43hoCeEnhPC8Uqmivb2BgGnrPyPLeW6J4CeF8Jg0XvSBDu7vnfNIbQ0keBQfrdNVxBWcw3n8EyV04HO4x+wpncG/IY3ndXTw5S/W1VgroAcPzmqv4C38HL/GEEYRFUYZHiZJtuPuGa7Vgbcf69PdzuWCzyLflnemkkpbauzfgFtr4P/AfrxcHTFrbpsp7xcDxFhEek1AHXj3QJ+RccqT1qepw4LRyYpNuDTlQIJ7a2z8I57GKURrbrPjvZ28txM+gnW48Mrp3KAkHRTjO1grhNeEsKsW3jXQZ2SMXM6n0mh/5PbAG0lQFqZT0I1PVM9/j69jcGrUOzIw3FF1ZRPe3LH+yJYPtV+6cPC3TzwixtuF8Adp+i+dy3igV9vRPkkQcjm9afRSjD4tG+XgyhvCyOVCvCYgXz3O4VsN4D2yxWlD9XotVl+ZWPFX2za/jbenHuz8ZZ9dla958Vhl1ei4R9PoSaysKehT716K6ZpbglCtgQSfxwTebAD/SQ0cyrKdcSBGjp5JdbVLxkq6KqmeGN0fo4cjn1Sz3gTeTRKb8JfK9n4t1cJKcbq2vhaAQ8tEyXdODKYbx4rGA63/GXcTemLUI1tR5+62wcDyTkOFiWyKtagTTcAFQiXqjdkhWjgCf0uCn14pSHdvyXlOnb2gBv6xWTm/3qiE4MC6nvCnfAvPtfwMjb8Jc9iN3mZ7X3D0wbFc4sdnzkfFLf3X7jcScKNsXVgq+G+S4LuV6P0PLpvZ1khACWNLBH8tFzwWo6Flbfz7S/0z2hsJeB+HZVNtsXE1BAdyiZ0p76zoZuTB/jkP1f0orRZiB/bgKdkiNefFQpHXB1NjxZnuheB0CPblc47HqFja2q9R1HXglVtfJkvBXuxrwoliYDgER5LEV1sSD6VjXr2xa354QwdmOdGJZ+s4kU6UHTn55/R3V4uGcomzrTl/L04qr70leGv9oabytOCPyTwiytiK41XHFhUL/htWO78qS8cPZTMEihhZNLlZAbNEfB8/wjheV7MDLjb+CxzUqEYXF2HCAAAAJXRFWHRkYXRlOmNyZWF0ZQAyMDIzLTA0LTIxVDA4OjAxOjAzKzAzOjAwMETKJgAAACV0RVh0ZGF0ZTptb2RpZnkAMjAyMy0wNC0yMVQwODowMTowMyswMzowMEEZcpoAAAAZdEVYdFNvZnR3YXJlAHd3dy5pbmtzY2FwZS5vcmeb7jwaAAAAAElFTkSuQmCC"),
        // Please specify the base64 content of a 80x80 pixels image
        ExportMetadata("BigImageBase64", "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"),
        ExportMetadata("BackgroundColor", "Lavender"),
        ExportMetadata("PrimaryFontColor", "Black"),
        ExportMetadata("SecondaryFontColor", "Gray")]
    public class MyPlugin : PluginBase
    {
        public override IXrmToolBoxPluginControl GetControl()
        {
            return new MyPluginControl();
        }

        /// <summary>
        /// Constructor 
        /// </summary>
        public MyPlugin()
        {
            // If you have external assemblies that you need to load, uncomment the following to 
            // hook into the event that will fire when an Assembly fails to resolve
            // AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AssemblyResolveEventHandler);
        }

        /// <summary>
        /// Event fired by CLR when an assembly reference fails to load
        /// Assumes that related assemblies will be loaded from a subfolder named the same as the Plugin
        /// For example, a folder named Sample.XrmToolBox.MyPlugin 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        private Assembly AssemblyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            Assembly loadAssembly = null;
            Assembly currAssembly = Assembly.GetExecutingAssembly();

            // base name of the assembly that failed to resolve
            var argName = args.Name.Substring(0, args.Name.IndexOf(","));

            // check to see if the failing assembly is one that we reference.
            List<AssemblyName> refAssemblies = currAssembly.GetReferencedAssemblies().ToList();
            var refAssembly = refAssemblies.Where(a => a.Name == argName).FirstOrDefault();

            // if the current unresolved assembly is referenced by our plugin, attempt to load
            if (refAssembly != null)
            {
                // load from the path to this plugin assembly, not host executable
                string dir = Path.GetDirectoryName(currAssembly.Location).ToLower();
                string folder = Path.GetFileNameWithoutExtension(currAssembly.Location);
                dir = Path.Combine(dir, folder);

                var assmbPath = Path.Combine(dir, $"{argName}.dll");

                if (File.Exists(assmbPath))
                {
                    loadAssembly = Assembly.LoadFrom(assmbPath);
                }
                else
                {
                    throw new FileNotFoundException($"Unable to locate dependency: {assmbPath}");
                }
            }

            return loadAssembly;
        }
    }
}