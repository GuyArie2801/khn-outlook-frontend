import win32com.client
import tkinter as tk
from tkinter import simpledialog, messagebox


def get_outlook():
    """Start Outlook COM connection."""
    app = win32com.client.Dispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")
    return app, ns   # return BOTH


def create_draft(app, ns, subject, body, recipient):
    """Creates a new draft safely in the user's main mailbox."""
    
    # Try multiple methods to create the draft
    methods = [
        ("Method 1: Direct Drafts folder creation", create_via_drafts_folder),
        ("Method 2: Inbox store creation", create_via_inbox_store),
        ("Method 3: First non-public store", create_via_first_store),
    ]
    
    for method_name, method_func in methods:
        try:
            print(f"\nTrying {method_name}...")
            result = method_func(app, ns, subject, body, recipient)
            if result:
                return True
        except Exception as e:
            print(f"{method_name} failed: {e}")
            import traceback
            traceback.print_exc()
    
    print("\nAll methods failed. Please check:")
    print("1. Outlook is running in normal mode (not safe mode)")
    print("2. You have a default email account configured")
    print("3. Your account has permission to create drafts")
    return False


def create_via_drafts_folder(app, ns, subject, body, recipient):
    """Try creating via drafts folder Items.Add"""
    inbox = ns.GetDefaultFolder(6)
    store = inbox.Store
    
    # Find drafts in this store
    drafts = None
    root = store.GetRootFolder()
    for folder in root.Folders:
        print(f"  Checking folder: {folder.Name}")
        if folder.DefaultItemType == 0:  # Mail folders
            for subfolder in folder.Folders:
                if subfolder.DefaultMessageClass == "IPM.Note" and "draft" in subfolder.Name.lower():
                    drafts = subfolder
                    break
        if drafts:
            break
    
    if not drafts:
        # Try GetDefaultFolder as fallback
        drafts = ns.GetDefaultFolder(16)
    
    print(f"  Found drafts at: {drafts.FolderPath}")
    
    # Create using Application with explicit class
    mail = app.CreateItem(0)
    mail.To = recipient
    mail.Subject = "Re: " + subject
    mail.HTMLBody = f"<p>Hello,</p><p>This is a generated draft:</p><hr>{body}"
    mail.Save()
    
    print(f"  Draft created successfully!")
    return True


def create_via_inbox_store(app, ns, subject, body, recipient):
    """Try creating via inbox store"""
    # Get inbox to find the right store
    inbox = ns.GetDefaultFolder(6)
    print(f"  Using store: {inbox.Store.DisplayName}")
    
    # Try to use Application.Session to create item
    session = app.Session
    mail = session.CreateItem(0)
    
    mail.To = recipient
    mail.Subject = "Re: " + subject  
    mail.HTMLBody = f"<p>Hello,</p><p>This is a generated draft:</p><hr>{body}"
    mail.Save()
    
    print(f"  Draft created via session!")
    return True


def create_via_first_store(app, ns, subject, body, recipient):
    """Try using the first non-public store"""
    for store in ns.Stores:
        print(f"  Checking store: {store.DisplayName} (Type: {store.ExchangeStoreType})")
        # Skip public folders (type 2) and archive stores (type 1)
        if store.ExchangeStoreType not in [1, 2]:
            print(f"  Using store: {store.DisplayName}")
            
            # Get root and find a mail folder
            root = store.GetRootFolder()
            
            # Create item
            mail = app.CreateItem(0)
            mail.To = recipient
            mail.Subject = "Re: " + subject
            mail.HTMLBody = f"<p>Hello,</p><p>This is a generated draft:</p><hr>{body}"
            mail.Save()
            
            print(f"  Draft created in primary store!")
            return True
    
    raise Exception("No suitable store found")


def find_public_folder(root, name):
    """Recursively search for a folder by name."""
    for folder in root.Folders:
        if folder.Name.lower() == name.lower():
            return folder
        found = find_public_folder(folder, name)
        if found:
            return found
    return None


def select_email_gui(email_subjects):
    root = tk.Tk()
    root.withdraw()

    choice = simpledialog.askinteger(
        "Select Email",
        "\n".join(f"{i+1}. {s}" for i, s in enumerate(email_subjects)) +
        "\n\nEnter the number of the email to process:",
        minvalue=1,
        maxvalue=len(email_subjects)
    )

    return None if choice is None else choice - 1


def main():
    app, ns = get_outlook()  # Fixed: clearer naming

    # Find public folder
    public_store = None
    for store in ns.Stores:
        if store.ExchangeStoreType == 2:
            public_store = store
            break

    if not public_store:
        messagebox.showerror("Error", "Could not find Public Folders store.")
        return

    root_folder = public_store.GetRootFolder()
    fax_folder = find_public_folder(root_folder, "fax")

    if not fax_folder:
        messagebox.showerror("Error", "Could not find the Fax folder.")
        return

    print(f"Found Fax folder at: {fax_folder.FolderPath}")

    items = fax_folder.Items
    if len(items) == 0:
        messagebox.showinfo("Info", "Fax folder is empty.")
        return

    email_subjects = [item.Subject for item in items]
    idx = select_email_gui(email_subjects)
    if idx is None:
        return

    selected_item = items[idx]

    print(f"Selected: {selected_item.Subject}")

    try:
        body = selected_item.HTMLBody
    except:
        body = selected_item.Body

    res = create_draft(
        app,
        ns,  # Pass both app and ns
        selected_item.Subject,
        body,
        "noa@khn.co.il"
    )

    if res:
        messagebox.showinfo("Success", "Draft created in Drafts folder!")


if __name__ == "__main__":
    main()