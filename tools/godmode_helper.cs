using System;

class GodModeHelper
{
    [STAThread]
    static int Main(string[] args)
    {
        try
        {
            Type shellType = Type.GetTypeFromProgID("Shell.Application");
            if (shellType == null)
            {
                Console.Error.WriteLine("Shell.Application not available.");
                return 3;
            }

            dynamic shell = Activator.CreateInstance(shellType);
            dynamic folder = shell.NameSpace("shell:::{ED7BA470-8E54-465E-825C-99712043E01C}");
            if (folder == null)
            {
                Console.Error.WriteLine("God Mode folder not available.");
                return 4;
            }

            if (args.Length > 0 && args[0] == "--invoke")
            {
                if (args.Length < 2)
                {
                    Console.Error.WriteLine("Missing item name for invoke.");
                    return 5;
                }

                string target = args[1];
                dynamic itemsForInvoke = folder.Items();
                foreach (dynamic entry in itemsForInvoke)
                {
                    string entryName = entry == null ? null : entry.Name as string;
                    if (string.IsNullOrWhiteSpace(entryName))
                    {
                        continue;
                    }
                    if (string.Equals(entryName, target, StringComparison.OrdinalIgnoreCase))
                    {
                        entry.InvokeVerb();
                        return 0;
                    }
                }

                Console.Error.WriteLine("Item not found: " + target);
                return 6;
            }

            dynamic items = folder.Items();
            foreach (dynamic item in items)
            {
                string name = item == null ? null : item.Name as string;
                if (!string.IsNullOrWhiteSpace(name))
                {
                    Console.WriteLine(name);
                }
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.ToString());
            return 1;
        }
    }
}
