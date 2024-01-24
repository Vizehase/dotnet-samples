# Goal

We have an old large 32bit C++ application (created with Borland C++ Builder 5) on its memory limits. We would like to have it call modules written in C# that reside in another process. With .NET 6, Microsoft propagated .NET COM servers and registry-free (!) access from a client, but the examples show a .NET client with in-process access only.

In short shown as a table:

| | |
| --- | --- |
| Client | 32bit native application |
| Server | .NET 8 DLL or EXE <br/> out-of-process <br/> registry-free |


# Relevant Microsoft demos

* [Reg-free, but .NET client, in-proc](https://github.com/dotnet/samples/tree/main/core/extensions/COMServerDemo)
* [Native client/.NET server, out-of-proc, but not reg-free](https://learn.microsoft.com/en-us/samples/dotnet/samples/out-of-process-com-server/)

# Out-of-Proc

The out-of-process part seems to come for free anyway: When I register (``regsvr32``) my .NET dll (built for target AnyCPU), it can be called from my 32bit native app. Because AnyCPU means 64bit for COM on a 64bit Windows, it automatically starts the built-in surrogate process dllhost.exe. So our remaining task is the registry-free configuration with application manifests.

# Actions taken

We take the [Microsoft out-of-proc example](https://learn.microsoft.com/en-us/samples/dotnet/samples/out-of-process-com-server/) as base with the following changes:

1. Make the DllServer project registry-free by adding ``<EnableRegFreeCom>True</EnableRegFreeCom>`` next to the EnableComHosting tag.
1. Open the Solution with Visual Studio (at the time of writing VS2022).
1. Open the Configuration Manager, create a new Solution Platform "x86" on base of the existing, leave all the .NET projects with platform "Any CPU", but change the C++ projects NativeClient and Server.Contract to Win32.
1. Compile the DllServer project and copy its output to the NativeClient's output directory (``..\Debug\`` relative to the project directory).
1. Create a new file ``NativeClient.manifest`` in the NativeClient project directory with the following contents (with or without the ``assemblyIdentity`` for NativeClient does not seem to make a difference although the tag is officially required in a manifest):

    ``` xml
    <?xml version="1.0" encoding="utf-8"?>
    <!-- https://docs.microsoft.com/windows/desktop/sbscs/assembly-manifests -->
    <assembly manifestVersion="1.0" xmlns="urn:schemas-microsoft-com:asm.v1">
      <!-- <assemblyIdentity
        type="win32" 
        name="NativeClient"
        version="1.0.0.0" /> -->

      <dependency>
        <dependentAssembly>
          <!-- RegFree COM matching the registration of the generated managed COM server -->
          <assemblyIdentity
              type="win32"
              name="DllServer.X"
              version="1.0.0.0"/>
        </dependentAssembly>
      </dependency>
    </assembly>
    ```

1. Compile the NativeClient project.
1. Embed manifest in client EXE: Open a Visual Studio Command Prompt, cd to the NativeClient directory and execute:

    ```console
    mt.exe -manifest NativeClient.manifest -outputresource:..\Debug\NativeClient.exe;1
    ```

1. Run the Client. For me, it says ``CoCreateInstance failure: 0x80040154`` (REGDB_E_CLASSNOTREG) <br/>Why? What is the correct application manifest?


# sxstrace

I was running

    sxstrace trace -logfile:sxstrace.etl

while starting ``NativeClient.exe``. When the client had exited, I stopped sxstrace and translated the log with

    sxstrace parse -logfile:sxstrace.etl -outfile:sxstrace.txt

but the NativeClient did not show up here.