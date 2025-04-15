using System;
using Nuke.Common;
using Nuke.Common.CI.GitHubActions;
using Nuke.Common.Tools.DotNet;
using Nuke.Common.IO;
using Nuke.Common.Utilities.Collections;
using static Nuke.Common.Tools.DotNet.DotNetTasks;

[GitHubActions(
    "deploy-to-nuget",
    GitHubActionsImage.UbuntuLatest,
    OnPushTags = new[] { "v*" },
    InvokedTargets = new[] { nameof(Pack), nameof(Publish) },
    ImportSecrets = new[] { "NUGET_API_KEY" },
    AutoGenerate = false)]
class Build : NukeBuild
{
    public static int Main() => Execute<Build>(x => x.Pack);

    [Parameter] readonly string Configuration = "Release";
    [Parameter][Secret] readonly string NugetApiKey;
    [Parameter] readonly string Version;

    AbsolutePath SourceDirectory => RootDirectory / "src";
    AbsolutePath ArtifactsDirectory => RootDirectory / "artifacts";

    Target Clean => _ => _
        .Executes(() =>
        {
            SourceDirectory.GlobDirectories("**/bin", "**/obj").DeleteDirectories();
            ArtifactsDirectory.CreateOrCleanDirectory();
        });

    Target Restore => _ => _
        .DependsOn(Clean)
        .Executes(() =>
        {
            SourceDirectory.GlobFiles(SourceDirectory, "**/*.csproj")
                .ForEach(project => DotNetRestore(s => s
                    .SetProjectFile(project)));
        });

    Target Compile => _ => _
        .DependsOn(Restore)
        .Executes(() =>
        {
            SourceDirectory.GlobFiles(SourceDirectory, "**/*.csproj")
                .ForEach(project => DotNetBuild(s => s
                    .SetProjectFile(project)
                    .SetConfiguration(Configuration)
                    .SetVersion(Version)
                    .EnableNoRestore()));
        });

    Target Pack => _ => _
        .DependsOn(Compile)
        .Executes(() =>
        {
            SourceDirectory.GlobFiles(SourceDirectory, "**/*.csproj")
                .ForEach(project => DotNetPack(s => s
                   .SetProject(project)
                   .SetConfiguration(Configuration)
                   .SetVersion(Version)
                   .SetOutputDirectory(ArtifactsDirectory)));

        });

    Target Publish => _ => _
        .DependsOn(Pack)
        .Requires(() => NugetApiKey)
        .Executes(() =>
        {
            DotNetNuGetPush(s => s
                .SetTargetPath(ArtifactsDirectory / "*.nupkg")
                .SetSource("https://api.nuget.org/v3/index.json")
                .SetApiKey(NugetApiKey));
        });
}