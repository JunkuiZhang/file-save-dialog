use windows::{
    core::*,
    System::{DispatcherQueue, DispatcherQueueHandler},
    Win32::{
        Foundation::*,
        Graphics::Gdi::ValidateRect,
        System::{
            Com::*,
            LibraryLoader::GetModuleHandleA,
            Ole::OleInitialize,
            WinRT::{
                CreateDispatcherQueueController, DispatcherQueueOptions, DQTAT_COM_NONE,
                DQTYPE_THREAD_CURRENT,
            },
        },
        UI::{
            Shell::{Common::*, *},
            WindowsAndMessaging::*,
        },
    },
};

fn open_dialog() -> Result<()> {
    unsafe {
        // CoIncrementMTAUsage()?;

        let dialog: IFileSaveDialog = CoCreateInstance(&FileSaveDialog, None, CLSCTX_ALL)?;

        dialog.SetFileTypes(&[
            COMDLG_FILTERSPEC {
                pszName: w!("Text files"),
                pszSpec: w!("*.txt"),
            },
            COMDLG_FILTERSPEC {
                pszName: w!("All files"),
                pszSpec: w!("*.*"),
            },
        ])?;

        if dialog.Show(None).is_ok() {
            let result = dialog.GetResult()?;
            let path = result.GetDisplayName(SIGDN_FILESYSPATH)?;
            println!("user picked: {}", path.display());
            CoTaskMemFree(Some(path.0 as _));
        } else {
            println!("user canceled");
        }

        Ok(())
    }
}

fn main() -> Result<()> {
    unsafe {
        OleInitialize(None)?;

        let options = DispatcherQueueOptions {
            dwSize: std::mem::size_of::<DispatcherQueueOptions>() as u32,
            threadType: DQTYPE_THREAD_CURRENT,
            apartmentType: DQTAT_COM_NONE,
        };
        let con = CreateDispatcherQueueController(options).unwrap();
        con.DispatcherQueue().unwrap();

        let instance = GetModuleHandleA(None)?;
        let window_class = s!("window");

        let wc = WNDCLASSA {
            hCursor: LoadCursorW(None, IDC_ARROW)?,
            hInstance: instance.into(),
            lpszClassName: window_class,

            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(wndproc),
            ..Default::default()
        };

        let atom = RegisterClassA(&wc);
        debug_assert!(atom != 0);

        CreateWindowExA(
            WINDOW_EX_STYLE::default(),
            window_class,
            s!("This is a sample window"),
            WS_OVERLAPPEDWINDOW | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            None,
            None,
            instance,
            None,
        )?;

        let mut message = MSG::default();

        while GetMessageA(&mut message, None, 0, 0).into() {
            DispatchMessageA(&message);
        }

        Ok(())
    }
}

extern "system" fn wndproc(window: HWND, message: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        match message {
            WM_PAINT => {
                println!("WM_PAINT");
                _ = ValidateRect(window, None);
                LRESULT(0)
            }
            WM_KEYDOWN => {
                let handler = DispatcherQueueHandler::new(move || {
                    open_dialog().unwrap();
                    Ok(())
                });
                let queue = DispatcherQueue::GetForCurrentThread().unwrap();
                queue.TryEnqueue(&handler).unwrap();
                LRESULT(0)
            }
            WM_DESTROY => {
                println!("WM_DESTROY");
                PostQuitMessage(0);
                LRESULT(0)
            }
            _ => DefWindowProcA(window, message, wparam, lparam),
        }
    }
}
