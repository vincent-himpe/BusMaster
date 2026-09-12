' ============================================================================
'  Modules\STM32DFU.vb
'
'  Talking to an STM32 sitting in its DFU bootloader.
'
'  This is the only place that knows ST's command line programmer exists. Above
'  it, ProbeFirmware deals in three questions - what is out there, put this file
'  on it, and get the running probe into the bootloader - and nothing about how
'  those are answered leaks upwards. When the programmer is called something else,
'  or is replaced by a library, this file changes and nothing else does.
'
'  **The bodies are not written yet.** The signatures and the answers they give
'  are settled, because the window above is built on them; what goes inside is
'  coming from the user's own sample of driving the programmer.
'
'  Every one of them returns a string rather than throwing, because every one of
'  them is a conversation with a lump of hardware over a cable that may not be
'  plugged in. "Nothing found" is an ordinary answer here, not an exception. That
'  goes for the programmer itself, which may simply not be installed.
' ============================================================================

Imports System.IO

Public Module STM32DFU

    ''' <summary>
    ''' What FlashFile says when the part took the firmware, and when it did not.
    '''
    ''' Constants rather than literals at both ends, so the two sides of the answer
    ''' cannot drift apart. Anything that is not Succeeded is treated as a failure
    ''' by the caller - with a part that can be bricked, the only safe reading of an
    ''' answer nobody recognises is that it did not work.
    ''' </summary>
    Public Const Succeeded As String = "SUCCESS"
    Public Const Failed As String = "FAIL"

    ''' <summary>
    ''' ST's command line programmer. The one name, used both when walking PATH and
    ''' when looking in the places the installer puts it.
    ''' </summary>
    Private Const ProgrammerExe As String = "STM32_Programmer_CLI.exe"

    ''' <summary>
    ''' Where FindProgrammer last found the programmer, or an empty string.
    '''
    ''' Private on purpose: nothing outside this file has any business knowing that
    ''' the work is done by running an executable, let alone which one. The two
    ''' functions below are what need it.
    ''' </summary>
    Private ProgrammerPath As String = String.Empty


    ''' <summary>
    ''' Finds STM32CubeProgrammer's command line executable, and remembers where it
    ''' is. Returns the full path, or an empty string when it is not installed.
    '''
    ''' PATH first, because someone who has put it there meant that one to be used;
    ''' then the two folders the installer uses, which is where it is on a machine
    ''' nobody has arranged. Looked for afresh on every scan rather than once at
    ''' start-up, so installing it while the window is open is enough.
    '''
    ''' An empty string rather than Nothing for "not found", so it reads the same
    ''' way as everything else here. String.IsNullOrEmpty cannot tell the two apart
    ''' anyway, so nothing that checks it needs to care which it got.
    ''' </summary>
    Public Function FindProgrammer() As String

        ProgrammerPath = String.Empty

        ' 1. Anywhere on PATH.
        Dim pathVariable As String = Environment.GetEnvironmentVariable("PATH")

        If Not String.IsNullOrEmpty(pathVariable) Then

            For Each pathDir As String In pathVariable.Split(";"c)

                Dim folder As String = pathDir.Trim()
                If folder.Length = 0 Then Continue For

                ' A PATH with a malformed entry in it is common and is not this
                ' program's problem; Path.Combine throws on one, and a bad entry
                ' halfway along must not stop the good ones being looked at.
                Try
                    Dim onPath As String = Path.Combine(folder, ProgrammerExe)

                    If File.Exists(onPath) Then
                        ProgrammerPath = onPath
                        Return ProgrammerPath
                    End If
                Catch
                    Continue For
                End Try

            Next

        End If

        ' 2. Where the installer puts it.
        Dim programFiles As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
        Dim programFilesX86 As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)

        Dim commonPaths As String() = {
            Path.Combine(programFiles, "STMicroelectronics", "STM32Cube", "STM32CubeProgrammer", "bin"),
            Path.Combine(programFilesX86, "STMicroelectronics", "STM32Cube", "STM32CubeProgrammer", "bin")
        }

        For Each folder As String In commonPaths

            Dim installed As String = Path.Combine(folder, ProgrammerExe)

            If File.Exists(installed) Then
                ProgrammerPath = installed
                Return ProgrammerPath
            End If

        Next

        Return String.Empty

    End Function


    ''' <summary>
    ''' The serial number of the STM32 waiting in DFU mode, or an empty string when
    ''' there is not one.
    '''
    ''' Empty is the everyday answer, not a fault: it is what comes back when the
    ''' BOOT and RESET dance did not take, which is most first attempts.
    ''' </summary>
    Public Function GetDFUSerialNumber() As String

        ' TODO: run ProgrammerPath with its list-devices switch, and pick the serial
        ' out of what it prints. Sample to come from the user. FindProgrammer has
        ' already been called by the time the scan gets here, so ProgrammerPath is
        ' set - but check it rather than trusting the caller.
        Return String.Empty

    End Function

    ''' <summary>
    ''' Writes a firmware image to the part that is in DFU mode. Returns Succeeded
    ''' or Failed.
    '''
    ''' The part reboots out of the bootloader afterwards, so the serial number the
    ''' scan found is stale the moment this returns, whichever way it went.
    ''' </summary>
    Public Function FlashFile(firmwareFile As String) As String

        ' TODO: run ProgrammerPath against firmwareFile, wait for it, and read
        ' success or failure out of its exit code and output. Call FindProgrammer
        ' first if ProgrammerPath is empty - nothing guarantees a scan came before.
        Return Failed

    End Function

    ''' <summary>
    ''' Asks a probe that is already running to reboot into its bootloader, so it
    ''' can be reflashed without anyone holding BOOT down. Returns Succeeded or
    ''' Failed.
    '''
    ''' This is the one that makes Update Firmware different from Provision Probe:
    ''' the provisioning case has no firmware on the part to ask.
    ''' </summary>
    Public Function EnterBootloader(targetDevice As String) As String

        ' TODO: send the running probe whatever command puts it in DFU, then wait
        ' for it to come back on USB as a DFU device. Needs ProgrammerPath for the
        ' second half, so the same rule applies as in FlashFile.
        Return Failed

    End Function

End Module
