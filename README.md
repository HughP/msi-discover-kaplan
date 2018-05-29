# msi-discover-kaplan
This is a windows script for discovering data from a keyboard layout written by Michael Kaplan

```
using System;
using System.Text;
using System.Collections;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace KeyboardLayouts {

    //  You'll want to insert that enumeration from part #0 here!

    public enum ShiftState : int {
        Base            = 0,                    // 0
        Shft            = 1,                    // 1
        Ctrl            = 2,                    // 2
        ShftCtrl        = Shft | Ctrl,          // 3
        Menu            = 4,                    // 4 -- NOT USED
        ShftMenu        = Shft | Menu,          // 5 -- NOT USED
        MenuCtrl        = Menu | Ctrl,          // 6
        ShftMenuCtrl    = Shft | Menu | Ctrl,   // 7
        Xxxx            = 8,                    // 8
        ShftXxxx        = Shft | Xxxx,          // 9
    }

    public class DeadKey {
        private char m_deadchar;
        private ArrayList m_rgbasechar = new ArrayList();
        private ArrayList m_rgcombchar = new ArrayList();

        public DeadKey(char deadCharacter) {
            this.m_deadchar = deadCharacter;
        }

        public char DeadCharacter {
            get {
                return this.m_deadchar;
            }
        }

        public void AddDeadKeyRow(char baseCharacter, char combinedCharacter) {
            this.m_rgbasechar.Add(baseCharacter);
            this.m_rgcombchar.Add(combinedCharacter);
        }

        public int Count {
            get {
                return this.m_rgbasechar.Count;
            }
        }

        public char GetBaseCharacter(int index) {
            return (char)this.m_rgbasechar[index];
        }

        public char GetCombinedCharacter(int index) {
            return (char)this.m_rgcombchar[index];
        }

        public bool ContainsBaseCharacter(char baseCharacter) {
            return this.m_rgbasechar.Contains(baseCharacter);
        }
    }

    public class VirtualKey {
        [DllImport("user32.dll", CharSet=CharSet.Unicode, EntryPoint="MapVirtualKeyExW", ExactSpelling=true)]
        internal static extern uint MapVirtualKeyEx(
            uint uCode,
            uint uMapType,
            IntPtr dwhkl);

        private IntPtr m_hkl;
        private uint m_vk;
        private uint m_sc;
        private bool[,] m_rgfDeadKey = new bool[(int)ShiftState.ShftXxxx + 1,2];
        private string[,] m_rgss = new string[(int)ShiftState.ShftXxxx + 1,2];

        public VirtualKey(IntPtr hkl, KeysEx virtualKey) {
            this.m_sc = MapVirtualKeyEx((uint)virtualKey, 0, hkl);
            this.m_hkl = hkl;
            this.m_vk = (uint)virtualKey;
        }

        public VirtualKey(IntPtr hkl, uint scanCode) {
            this.m_vk = MapVirtualKeyEx(scanCode, 1, hkl);
            this.m_hkl = hkl;
            this.m_sc = scanCode;
        }

        public KeysEx VK {
            get { return (KeysEx)this.m_vk; }
        }

        public uint SC {
            get { return this.m_sc; }
        }

        public string GetShiftState(ShiftState shiftState, bool capsLock) {
            if(this.m_rgss[(uint)shiftState, (capsLock ? 1 : 0)] == null) {
                return("");
            }
           
            return(this.m_rgss[(uint)shiftState, (capsLock ? 1 : 0)]);
        }

        public void SetShiftState(ShiftState shiftState, string value, bool isDeadKey, bool capsLock) {
            this.m_rgfDeadKey[(uint)shiftState, (capsLock ? 1 : 0)] = isDeadKey;
            this.m_rgss[(uint)shiftState, (capsLock ? 1 : 0)] = value;
        }

        public bool IsSGCAPS {
            get {
                string stBase = this.GetShiftState(ShiftState.Base, false);
                string stShift = this.GetShiftState(ShiftState.Shft, false);
                string stCaps = this.GetShiftState(ShiftState.Base, true);
                string stShiftCaps = this.GetShiftState(ShiftState.Shft, true);
                return(
                    ((stCaps.Length > 0) &&
                    (! stBase.Equals(stCaps)) &&
                    (! stShift.Equals(stCaps))) ||
                    ((stShiftCaps.Length > 0) &&
                    (! stBase.Equals(stShiftCaps)) &&
                    (! stShift.Equals(stShiftCaps))));
            }
        }

        public bool IsCapsEqualToShift {
            get {
                string stBase = this.GetShiftState(ShiftState.Base, false);
                string stShift = this.GetShiftState(ShiftState.Shft, false);
                string stCaps = this.GetShiftState(ShiftState.Base, true);
                return(
                    (stBase.Length > 0) &&
                    (stShift.Length > 0) &&
                    (! stBase.Equals(stShift)) &&
                    (stShift.Equals(stCaps)));
            }
        }

        public bool IsAltGrCapsEqualToAltGrShift {
            get {
                string stBase = this.GetShiftState(ShiftState.MenuCtrl, false);
                string stShift = this.GetShiftState(ShiftState.ShftMenuCtrl, false);
                string stCaps = this.GetShiftState(ShiftState.MenuCtrl, true);
                return(
                    (stBase.Length > 0) &&
                    (stShift.Length > 0) &&
                    (! stBase.Equals(stShift)) &&
                    (stShift.Equals(stCaps)));
            }
        }

        public bool IsXxxxGrCapsEqualToXxxxShift {
            get {
                string stBase = this.GetShiftState(ShiftState.Xxxx, false);
                string stShift = this.GetShiftState(ShiftState.ShftXxxx, false);
                string stCaps = this.GetShiftState(ShiftState.Xxxx, true);
                return(
                    (stBase.Length > 0) &&
                    (stShift.Length > 0) &&
                    (! stBase.Equals(stShift)) &&
                    (stShift.Equals(stCaps)));
            }
        }

        public bool IsEmpty {
            get {
                for(int i = 0; i < this.m_rgss.GetUpperBound(0); i++) {
                    for(int j = 0; j <= 1; j++) {
                        if(this.GetShiftState((ShiftState)i, (j == 1)).Length > 0) {
                            return(false);
                        }
                    }
                }
                return true;
            }
        }

        public string LayoutRow {
            get {
                StringBuilder sbRow = new StringBuilder();

                // First, get the SC/VK info stored
                sbRow.Append(string.Format("{0:x2}\t{1:x2} - {2}", this.SC, (byte)this.VK, ((KeysEx)this.VK).ToString().PadRight(13)));

                // Now the CAPSLOCK value
                int capslock =
                    0 |
                    (this.IsCapsEqualToShift ? 1 : 0) |
                    (this.IsSGCAPS ? 2 : 0) |
                    (this.IsAltGrCapsEqualToAltGrShift ? 4 : 0) |
                    (this.IsXxxxGrCapsEqualToXxxxShift ? 8 : 0);
                sbRow.Append(string.Format("\t{0}", capslock));


                for(ShiftState ss = 0; ss <= Loader.MaxShiftState; ss++) {
                    if(ss == ShiftState.Menu || ss == ShiftState.ShftMenu) {
                        // Alt and Shift+Alt don't work, so skip them
                        continue;
                    }
                    for(int caps = 0; caps <= 1; caps++) {
                        string st = this.GetShiftState(ss, (caps == 1));

                        if(st.Length == 0) {
                            // No character assigned here, put in -1.
                            sbRow.Append("\t  -1");
                        }
                        else if((caps == 1) && st == (this.GetShiftState(ss, (caps == 0)))) {
                            // Its a CAPS LOCK state and the assigned character(s) are
                            // identical to the non-CAPS LOCK state. Put in a MIDDLE DOT.
                            sbRow.Append("\t   \u00b7");
                        }
                        else if(this.m_rgfDeadKey[(int)ss, caps]) {
                            // It's a dead key, append an @ sign.
                            sbRow.Append(string.Format("\t{0:x4}@", ((ushort)st[0])));
                        }
                        else {
                            // It's some characters; put 'em in there.
                            StringBuilder sbChar = new StringBuilder((5 * st.Length) + 1);
                            for(int ich = 0; ich < st.Length; ich++) {
                                sbChar.Append(((ushort)st[ich]).ToString("x4"));
                                sbChar.Append(' ');
                            }
                            sbRow.Append(string.Format("\t{0}", sbChar.ToString(0, sbChar.Length - 1)));
                        }
                    }
                }

                return sbRow.ToString();
            }
        }
    }

    public class Loader {

        private const uint KLF_NOTELLSHELL  = 0x00000080;

        internal static KeysEx[] lpKeyStateNull = new KeysEx[256];

        [DllImport("user32.dll", CharSet=CharSet.Unicode, EntryPoint="LoadKeyboardLayoutW", ExactSpelling=true)]
        private static extern IntPtr LoadKeyboardLayout(string pwszKLID, uint Flags);

        [DllImport("user32.dll", ExactSpelling=true)]
        private static extern bool UnloadKeyboardLayout(IntPtr hkl);

        [DllImport("user32.dll", CharSet=CharSet.Unicode, ExactSpelling=true)]
        private static extern int ToUnicodeEx(
            uint wVirtKey,
            uint wScanCode,
            KeysEx[] lpKeyState,
            StringBuilder pwszBuff,
            int cchBuff,
            uint wFlags,
            IntPtr dwhkl);

        [DllImport("user32.dll", CharSet=CharSet.Unicode, EntryPoint="VkKeyScanExW", ExactSpelling=true)]
        private static extern ushort VkKeyScanEx(char ch, IntPtr dwhkl);

        [DllImport("user32.dll", ExactSpelling=true)]
        private static extern int GetKeyboardLayoutList(int nBuff, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] lpList);

        private static KeysEx m_XxxxVk = KeysEx.None;
        public static KeysEx XxxxVk {
            get {
                return m_XxxxVk;
            }
            set {
                m_XxxxVk = value;
            }
        }

        public static ShiftState MaxShiftState {
            get {
                return (Loader.XxxxVk == KeysEx.None ? ShiftState.ShftMenuCtrl : ShiftState.ShftXxxx);
            }
        }

        private static void FillKeyState(KeysEx[] lpKeyState, ShiftState ss, bool fCapsLock) {
            lpKeyState[(int)KeysEx.VK_SHIFT]    = (((ss & ShiftState.Shft) != 0) ? (KeysEx)0x80 : (KeysEx)0x00);
            lpKeyState[(int)KeysEx.VK_CONTROL]  = (((ss & ShiftState.Ctrl) != 0) ? (KeysEx)0x80 : (KeysEx)0x00);
            lpKeyState[(int)KeysEx.VK_MENU]     = (((ss & ShiftState.Menu) != 0) ? (KeysEx)0x80 : (KeysEx)0x00);
            if(Loader.XxxxVk != KeysEx.None) {
                // The Xxxx key has been assigned, so let's include it
                lpKeyState[(int)Loader.XxxxVk]       = (((ss & ShiftState.Xxxx) != 0) ? (KeysEx)0x80 : (KeysEx)0x00);
            }
            lpKeyState[(int)KeysEx.VK_CAPITAL]  = (fCapsLock ? (KeysEx)0x01 : (KeysEx)0x00);
        }

        private static DeadKey ProcessDeadKey(
            uint iKeyDead,              // The index into the VirtualKey of the dead key
            ShiftState shiftStateDead,  // The shiftstate that contains the dead key
            KeysEx[] lpKeyStateDead,    // The key state for the dead key
            VirtualKey[] rgKey,         // Our array of dead keys
            bool fCapsLock,             // Was the caps lock key pressed?
            IntPtr hkl) {               // The keyboard layout

            KeysEx[] lpKeyState = new KeysEx[256];
            DeadKey deadKey = new DeadKey(rgKey[iKeyDead].GetShiftState(shiftStateDead, fCapsLock)[0]);

            for(uint iKey = 0; iKey < rgKey.Length; iKey++) {
                if(rgKey[iKey] != null) {
                    StringBuilder sbBuffer = new StringBuilder(10);     // Scratchpad we use many places

                    for(ShiftState ss = ShiftState.Base; ss <= Loader.MaxShiftState; ss++) {
                        int rc = 0;
                        if(ss == ShiftState.Menu || ss == ShiftState.ShftMenu) {
                            // Alt and Shift+Alt don't work, so skip them
                            continue;
                        }

                        for(int caps = 0; caps <=1; caps++) {
                            // First the dead key
                            while(rc >= 0) {
                                // We know that this is a dead key coming up, otherwise
                                // this function would never have been called. If we do
                                // *not* get a dead key then that means the state is
                                // messed up so we run again and again to clear it up.
                                // Risk is technically an infinite loop but per Hiroyama
                                // that should be impossible here.
                                rc = ToUnicodeEx((uint)rgKey[iKeyDead].VK, rgKey[iKeyDead].SC, lpKeyStateDead, sbBuffer, sbBuffer.Capacity, 0, hkl);
                            }

                            // Now fill the key state for the potential base character
                            FillKeyState(lpKeyState, ss, (caps != 0));

                            sbBuffer = new StringBuilder(10);
                            rc = ToUnicodeEx((uint)rgKey[iKey].VK, rgKey[iKey].SC, lpKeyState, sbBuffer, sbBuffer.Capacity, 0, hkl);
                            if(rc == 1) {
                                // That was indeed a base character for our dead key.
                                // And we now have a composite character. Let's run
                                // through one more time to get the actual base
                                // character that made it all possible?
                                char combchar = sbBuffer[0];
                                sbBuffer = new StringBuilder(10);
                                rc = ToUnicodeEx((uint)rgKey[iKey].VK, rgKey[iKey].SC, lpKeyState, sbBuffer, sbBuffer.Capacity, 0, hkl);

                                char basechar = sbBuffer[0];

                                if(deadKey.DeadCharacter == combchar) {
                                    // Since the combined character is the same as the dead key,
                                    // we must clear out the keyboard buffer.
                                    ClearKeyboardBuffer((uint)KeysEx.VK_DECIMAL, rgKey[(uint)KeysEx.VK_DECIMAL].SC, hkl);
                                }

                                if((((ss == ShiftState.Ctrl) || (ss == ShiftState.ShftCtrl)) &&
                                    (char.IsControl(basechar))) ||
                                    (basechar.Equals(combchar))) {
                                    // ToUnicodeEx has an internal knowledge about those
                                    // VK_A ~ VK_Z keys to produce the control characters,
                                    // when the conversion rule is not provided in keyboard
                                    // layout files

                                    // Additionally, dead key state is lost for some of these
                                    // character combinations, for unknown reasons.

                                    // Therefore, if the base character and combining are equal,
                                    // and its a CTRL or CTRL+SHIFT state, and a control character
                                    // is returned, then we do not add this "dead key" (which
                                    // is not really a dead key).
                                    continue;
                                }

                                if(! deadKey.ContainsBaseCharacter(basechar)) {
                                    deadKey.AddDeadKeyRow(basechar, combchar);
                                }
                            }
                            else if(rc > 1) {
                                // Not a valid dead key combination, sorry! We just ignore it.
                            }
                            else if(rc < 0) {
                                // It's another dead key, so we ignore it (other than to flush it from the state)
                                ClearKeyboardBuffer((uint)KeysEx.VK_DECIMAL, rgKey[(uint)KeysEx.VK_DECIMAL].SC, hkl);
                            }
                        }
                    }
                }
            }
            return deadKey;
        }

        private static void ClearKeyboardBuffer(uint vk, uint sc, IntPtr hkl) {
            StringBuilder sb = new StringBuilder(10);
            int rc = 0;
            while(rc != 1) {
                rc = ToUnicodeEx(vk, sc, lpKeyStateNull, sb, sb.Capacity, 0, hkl);
            }
        }

        [STAThread]
        static void Main(string[] args) {
            int cKeyboards = GetKeyboardLayoutList(0, null);
            IntPtr[] rghkl = new IntPtr[cKeyboards];
            GetKeyboardLayoutList(cKeyboards, rghkl);           
            IntPtr hkl = LoadKeyboardLayout(args[0], KLF_NOTELLSHELL);
            if(hkl == IntPtr.Zero) {
                Console.WriteLine("Sorry, that keyboard does not seem to be valid.");
            }
            else {
                KeysEx[] lpKeyState = new KeysEx[256];
                VirtualKey[] rgKey = new VirtualKey[256];
                ArrayList alDead = new ArrayList();

                // Scroll through the Scan Code (SC) values and get the valid Virtual Key (VK)
                // values in it. Then, store the SC in each valid VK so it can act as both a
                // flag that the VK is valid, and it can store the SC value.
                for(uint sc = 0x01; sc <= 0x7f; sc++) {
                    VirtualKey key = new VirtualKey(hkl, sc);
                    if(key.VK != 0) {
                        rgKey[(uint)key.VK] = key;
                    }
                }

                // add the special keys that do not get added from the code above
                for(KeysEx ke = KeysEx.VK_NUMPAD0; ke <= KeysEx.VK_NUMPAD9; ke++) {
                    rgKey[(uint)ke] = new VirtualKey(hkl, ke);
                }
                rgKey[(uint)KeysEx.VK_DIVIDE] = new VirtualKey(hkl, KeysEx.VK_DIVIDE);
                rgKey[(uint)KeysEx.VK_CANCEL] = new VirtualKey(hkl, KeysEx.VK_CANCEL);
                rgKey[(uint)KeysEx.VK_DECIMAL] = new VirtualKey(hkl, KeysEx.VK_DECIMAL);

                // See if there is a special shift state added
                for(KeysEx vk = KeysEx.None; vk <= KeysEx.VK_OEM_CLEAR; vk++) {
                    uint sc = VirtualKey.MapVirtualKeyEx((uint)vk, 0, hkl);
                    uint vkL = VirtualKey.MapVirtualKeyEx(sc, 1, hkl);
                    uint vkR = VirtualKey.MapVirtualKeyEx(sc, 3, hkl);
                    if((vkL != vkR) &&
                        ((uint)vk != vkL)) {
                        switch(vk) {
                            case KeysEx.VK_LCONTROL:
                            case KeysEx.VK_RCONTROL:
                            case KeysEx.VK_LSHIFT:
                            case KeysEx.VK_RSHIFT:
                            case KeysEx.VK_LMENU:
                            case KeysEx.VK_RMENU:
                                break;

                            default:
                                Loader.XxxxVk = vk;
                                break;
                        }
                    }
                }

                for(uint iKey = 0; iKey < rgKey.Length; iKey++) {
                    if(rgKey[iKey] != null) {
                        StringBuilder sbBuffer;     // Scratchpad we use many places

                        for(ShiftState ss = ShiftState.Base; ss <= Loader.MaxShiftState; ss++) {
                            if(ss == ShiftState.Menu || ss == ShiftState.ShftMenu) {
                                // Alt and Shift+Alt don't work, so skip them
                                continue;
                            }

                            for(int caps = 0; caps <= 1; caps++) {
                                ClearKeyboardBuffer((uint)KeysEx.VK_DECIMAL, rgKey[(uint)KeysEx.VK_DECIMAL].SC, hkl);
                                FillKeyState(lpKeyState, ss, (caps != 0));
                                sbBuffer = new StringBuilder(10)
                                int rc = ToUnicodeEx((uint)rgKey[iKey].VK, rgKey[iKey].SC, lpKeyState, sbBuffer, sbBuffer.Capacity, 0, hkl);
                                if(rc > 0) {
                                    if(sbBuffer.Length == 0) {
                                        // Someone defined NULL on the keyboard; let's coddle them
                                        rgKey[iKey].SetShiftState(ss, "\u0000", false, (caps != 0));
                                    }
                                    else {
                                        if((rc == 1) &&
                                            (ss == ShiftState.Ctrl || ss == ShiftState.ShftCtrl) &&
                                            ((int)rgKey[iKey].VK == ((uint)sbBuffer[0] + 0x40))) {
                                            // ToUnicodeEx has an internal knowledge about those
                                            // VK_A ~ VK_Z keys to produce the control characters,
                                            // when the conversion rule is not provided in keyboard
                                            // layout files
                                            continue;
                                        }
                                        rgKey[iKey].SetShiftState(ss, sbBuffer.ToString().Substring(0, rc), false, (caps != 0));
                                    }
                                }
                                else if(rc < 0) {
                                    rgKey[iKey].SetShiftState(ss, sbBuffer.ToString().Substring(0, 1), true, (caps != 0));

                                    // It's a dead key; let's flush out whats stored in the keyboard state.
                                    ClearKeyboardBuffer((uint)KeysEx.VK_DECIMAL, rgKey[(uint)KeysEx.VK_DECIMAL].SC, hkl);
                                    DeadKey dk = null;
                                    for(int iDead = 0; iDead < alDead.Count; iDead++) {
                                        dk = (DeadKey)alDead[iDead];
                                        if(dk.DeadCharacter == rgKey[iKey].GetShiftState(ss, caps != 0)[0]) {
                                            break;
                                        }
                                        dk = null;
                                    }
                                    if(dk == null) {
                                        alDead.Add(ProcessDeadKey(iKey, ss, lpKeyState, rgKey, caps == 1, hkl));
                                    }
                                }
                            }
                        }
                    }
                }

                foreach(IntPtr i in rghkl) {
                    if(hkl == i) {
                        hkl = IntPtr.Zero;
                        break;
                    }
                }

                if(hkl != IntPtr.Zero) {
                    UnloadKeyboardLayout(hkl);
                }

                // Okay, now we can dump the layout
                Console.Write("\nSC\tVK  \t\t\tCAPS\t_\t_C\ts\tsC\tc\tcC\tsc\tscC\t\tca\tcaC\tsca\tscaC");
                if(Loader.XxxxVk != KeysEx.None) {
                    Console.Write("\tx\txC\tsx\tsxC");
                }
                Console.WriteLine();
                Console.Write("==\t==========\t\t====\t====\t====\t====\t====\t====\t====\t====\t====\t====\t====\t====\t====");
                if(Loader.XxxxVk != KeysEx.None) {
                    Console.Write("\t====\t====\t====\t====");
                }
                Console.WriteLine();

                for(uint iKey = 0; iKey < rgKey.Length; iKey++) {
                    if((rgKey[iKey] != null) &&
                        ( !rgKey[iKey].IsEmpty)) {
                        Console.WriteLine(rgKey[iKey].LayoutRow);
                    }
                }

                foreach(DeadKey dk in alDead) {
                    Console.WriteLine();
                    Console.WriteLine("0x{0:x4}\t{1}", ((ushort)dk.DeadCharacter).ToString("x4"), dk.Count);
                    for(int id = 0; id < dk.Count; id++) {
                        Console.WriteLine("\t0x{0:x4}\t0x{1:x4}",
                            ((ushort)dk.GetBaseCharacter(id)).ToString("x4"),
                            ((ushort)dk.GetCombinedCharacter(id)).ToString("x4"));
                    }
                }
                Console.WriteLine();

            }
        }
    }
}

```
