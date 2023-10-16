using Siemens.Engineering;
using Siemens.Engineering.Download;
using Siemens.Engineering.Download.Configurations;
using Siemens.Engineering.HW;
using Siemens.Simatic.Simulation.Runtime;
using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;

string projectFile = args.SingleOrDefault() ?? @"c:\projects\<path_to_project>\project.ap18";

// Necessary to have the UI visible in CI/CD interface, otherwise it won't do anything.
StartPlcSimAdvancedUserInterface();

IInstance instance = GetOrRegisterInstance("dummy");

DownloadAndStart(projectFile, "PN/IE", "Siemens PLCSIM Virtual Ethernet Adapter", "1 X1");

instance.Run();

void StartPlcSimAdvancedUserInterface()
{
    Process simAdvancedUI = new Process
    {
        StartInfo = new ProcessStartInfo(@"C:\Program Files (x86)\Siemens\Automation\PLCSIMADV\bin\Siemens.Simatic.PlcSim.Advanced.UserInterface.exe")
    };
    simAdvancedUI.Start();
}

IInstance GetOrRegisterInstance(string name)
{
    if (!SimulationRuntimeManager.RegisteredInstanceInfo.Any(x => x.Name == name))
    {
        Console.WriteLine($"Registering instance {name}...");
        instance = SimulationRuntimeManager.RegisterInstance(ECPUType.CPU1518, name);
        instance.CommunicationInterface = ECommunicationInterface.TCPIP;
        instance.PowerOn();
        Thread.Sleep(4000); // Power on needs some time...
        SIPSuite4 sipSuite = new SIPSuite4("192.168.0.1", "255.255.255.0", string.Empty);
        instance.SetIPSuite(0, sipSuite, true);
    }

    instance = SimulationRuntimeManager.CreateInterface(name);
    return instance;
}

void DownloadAndStart(string projectFile, string modeName, string adapterName, string slotName)
{
    TiaPortal tiaPortal = GetOrCreateTiaPortal(projectFile);
    Project project = GetProject(tiaPortal, projectFile);

    DeviceItem cpu = project
        .Devices
        .First()
        .DeviceItems
        .Single(t => t.Classification == DeviceItemClassifications.CPU);

    var provider = cpu.GetService<DownloadProvider>();
    var mode = provider.Configuration.Modes.Find(modeName);
    var @interface = mode.PcInterfaces.Single(x => x.Name.Equals(adapterName, StringComparison.OrdinalIgnoreCase));
    var target = @interface.TargetInterfaces.Single(x => x.Name.Equals(slotName, StringComparison.OrdinalIgnoreCase));

    // Call to this method will result in an exception when certificate
    var result = provider.Download(target, PreDownload, (_) => { }, DownloadOptions.Software);

    foreach (var message in result.Messages)
    {
        Console.WriteLine(message.Message);
    }
}

static TiaPortal GetOrCreateTiaPortal(string projectFile)
    => TiaPortal
        .GetProcesses()
        .FirstOrDefault(t => t.ProjectPath is not null && t.ProjectPath.FullName.Equals(projectFile, StringComparison.OrdinalIgnoreCase))?
        .Attach()
    ?? new TiaPortal(TiaPortalMode.WithUserInterface);

static Project GetProject(TiaPortal portal, string projectFile)
    => portal.Projects.FirstOrDefault(t => t.Path.FullName.Equals(projectFile, StringComparison.OrdinalIgnoreCase))
       ?? portal.Projects.Open(new System.IO.FileInfo(projectFile));

static void PreDownload(DownloadConfiguration downloadConfiguration)
{
    if (downloadConfiguration is StopModules sm)
    {
        sm.CurrentSelection = StopModulesSelections.StopAll;
    }
}