import tkinter as tk

def create_test_gui():
    def on_submit():
        if selected_option.get() != "":
            print(f"Chiusura finestra. Selezionato: {selected_option.get()}")
            root.destroy()
        else:
            print("⚠️ Nessuna opzione selezionata. La finestra resta aperta.")

    root = tk.Tk()
    root.title("Finestra di Test - Sondaggi")
    root.geometry("400x200")
    root.configure(bg="white")

    # Etichetta
    label = tk.Label(root, text="Seleziona un'opzione:", font=("Arial", 14), bg="white")
    label.pack(pady=10)

    # Radio buttons (nessuno selezionato)
    selected_option = tk.StringVar(value="")

    radio1 = tk.Radiobutton(root, text="Opzione 1", variable=selected_option,
                            value="Opzione 1", font=("Arial", 12), bg="white")
    radio1.pack()

    radio2 = tk.Radiobutton(root, text="Opzione 2", variable=selected_option,
                            value="Opzione 2", font=("Arial", 12), bg="white")
    radio2.pack()

    # Bottone invia
    button = tk.Button(root, text="Invia", command=on_submit, font=("Arial", 12), bg="#007ACC", fg="white")
    button.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    create_test_gui()
