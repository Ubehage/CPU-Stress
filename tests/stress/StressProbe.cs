using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;

public static class StressProbe
{
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr OpenFileMapping(uint access, bool inherit, string name);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr MapViewOfFile(IntPtr mapping, uint access, uint high, uint low, UIntPtr size);
    [DllImport("kernel32.dll")] public static extern bool UnmapViewOfFile(IntPtr address);
    [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr handle);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetProcessTimes(IntPtr process, out ulong created, out ulong exited, out ulong kernel, out ulong user);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr CreateJobObject(IntPtr attributes, string name);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool SetInformationJobObject(IntPtr job, int informationClass, byte[] information, uint length);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SendMessageTimeout(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam, uint flags, uint timeout, out UIntPtr result);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr window, uint message, UIntPtr wParam, IntPtr lParam);
    private delegate bool EnumCallback(IntPtr window, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumCallback callback, IntPtr parameter);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int length);

    public static IntPtr FindForm(uint process)
    {
        IntPtr found = IntPtr.Zero;
        EnumWindows(delegate(IntPtr window, IntPtr ignored) {
            uint owner;
            GetWindowThreadProcessId(window, out owner);
            if (owner != process) return true;
            StringBuilder name = new StringBuilder(256);
            GetClassName(window, name, name.Capacity);
            if (name.ToString().StartsWith("ThunderRT6Form", StringComparison.Ordinal)) {
                found = window;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public sealed class CpuSample
    {
        public double CpuMilliseconds;
        public double TimestampMilliseconds;
    }
    public static CpuSample SampleCpu(IntPtr process)
    {
        ulong created, exited, kernel, user;
        long before = Stopwatch.GetTimestamp();
        if (!GetProcessTimes(process, out created, out exited, out kernel, out user))
            throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        long after = Stopwatch.GetTimestamp();
        return new CpuSample {
            CpuMilliseconds = (kernel + user) / 10000.0,
            TimestampMilliseconds = ((double)before + after) * 500.0 / Stopwatch.Frequency
        };
    }

    public static int ReadSlot(IntPtr memory, int slot, int field)
    {
        if (slot < 0 || slot >= 512 || field < 0 || field >= 4) throw new ArgumentOutOfRangeException();
        return Marshal.ReadInt32(memory, slot * 16 + field * 4);
    }
    public static int ExistingOwner()
    {
        IntPtr mapping = OpenFileMapping(4, false, "Local\\UbeCPUStress");
        if (mapping == IntPtr.Zero) return 0;
        IntPtr memory = IntPtr.Zero;
        try {
            memory = MapViewOfFile(mapping, 4, 0, 0, new UIntPtr(8192));
            return memory == IntPtr.Zero ? -1 : ReadSlot(memory, 0, 0);
        } finally {
            if (memory != IntPtr.Zero) UnmapViewOfFile(memory);
            CloseHandle(mapping);
        }
    }
    public static void WriteSlot(IntPtr memory, int slot, int field, int value)
    {
        if (slot < 0 || slot >= 512 || field < 0 || field >= 4) throw new ArgumentOutOfRangeException();
        Marshal.WriteInt32(memory, slot * 16 + field * 4, value);
    }
}
