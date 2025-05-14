import os
import shutil
import threading
import subprocess
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import sys, os, tkinter as tk

class BackupApp(tb.Window):
    def __init__(self):
        super().__init__(title="Sauvegarde IT ClicOnLine", themename="flatly", size=(800, 700))
        # —— gestion du chemin d'accès à l'icône —— #
        if getattr(sys, 'frozen', False):
            # En .exe PyInstaller : ressources dans _MEIPASS
            base_path = sys._MEIPASS
        else:
            # En mode script normal
            base_path = os.path.dirname(__file__)

        ico_path = os.path.join(base_path, "mon_icone.ico")
        try:
            self.iconbitmap(ico_path)
        except Exception as e:
            self.log(f"Impossible de charger l'icône .ico : {e}")
        self.sources = []
        self.setup_ui()
        threading.Thread(target=self.planification_automatique, daemon=True).start()

    # ----------------- Utilitaires ----------------- #
    def log(self, msg):
        with open("sauvegarde_log.txt", "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} - {msg}\n")

    def verifier_espace_disque(self, root_path):
        drive = os.path.splitdrive(root_path)[0] + os.sep
        return shutil.disk_usage(drive).free

    def verifier_destination_valide(self, src, dst):
        abs_src = os.path.abspath(src)
        abs_dst = os.path.abspath(dst)
        if os.path.splitdrive(abs_src)[0].lower() != os.path.splitdrive(abs_dst)[0].lower():
            return True
        try:
            if os.path.commonpath([abs_dst, abs_src]) == abs_src:
                messagebox.showerror(
                    "Erreur critique",
                    f"La destination '{dst}' est incluse dans la source '{src}' — sauvegarde annulée."
                )
                return False
        except ValueError:
            return True
        return True

    def compter_fichiers(self):
        count = 0
        for src in self.sources:
            for _, _, fichiers in os.walk(src):
                count += len(fichiers)
        return count

    def connecter_nas(self, chemin, utilisateur, mdp):
        if utilisateur and mdp:
            cmd = ['net', 'use', chemin, f'/user:{utilisateur}', mdp]
            subprocess.run(cmd, shell=True)

    def sauvegarder_image_systeme(self, destination, partition="C:"):
        try:
            cmd = f'wbadmin start backup -backupTarget:"{destination}" -include:{partition} -allCritical -quiet'
            subprocess.run(cmd, shell=True)
            self.log("Image système terminée")
        except Exception as e:
            self.log(f"Erreur sauvegarde système : {e}")
            messagebox.showerror("Image système", str(e))

    def copier_incrementiel(self, src, dst, total, compteur):
        if not os.path.exists(dst):
            os.makedirs(dst)
        for dossier, _, fichiers in os.walk(src):
            rel = os.path.relpath(dossier, src)
            dest_folder = os.path.join(dst, rel)
            os.makedirs(dest_folder, exist_ok=True)
            for f in fichiers:
                src_file = os.path.join(dossier, f)
                dst_file = os.path.join(dest_folder, f)
                try:
                    size = os.path.getsize(src_file)
                    if self.verifier_espace_disque(dst) < size + 50*1024*1024:
                        raise Exception("Espace disque insuffisant.")
                    if not os.path.exists(dst_file) or os.path.getmtime(src_file) > os.path.getmtime(dst_file):
                        shutil.copy2(src_file, dst_file)
                        self.log(f"Copié: {src_file} → {dst_file}")
                except Exception as e:
                    self.log(f"Erreur copie {src_file} : {e}")
                finally:
                    compteur[0] += 1
                    pourc = min(100, int((compteur[0]/total[0]) * 100))
                    self.pb['value'] = pourc
                    self.lbl_status.config(text=f"Progression: {compteur[0]}/{total[0]} ({pourc}%)")
                    self.update()

    def creer_snapshot_hardlink(self, src, dst):
        if os.path.exists(dst):
            shutil.rmtree(dst)
        for dossier, _, fichiers in os.walk(src):
            rel = os.path.relpath(dossier, src)
            dest_folder = os.path.join(dst, rel)
            os.makedirs(dest_folder, exist_ok=True)
            for f in fichiers:
                src_file = os.path.join(dossier, f)
                dst_file = os.path.join(dest_folder, f)
                try:
                    os.link(src_file, dst_file)
                except Exception:
                    shutil.copy2(src_file, dst_file)

    # ----------------- UI Setup ----------------- #
    def setup_ui(self):
        pad = dict(padx=10, pady=5)
        frm = tb.Frame(self, padding=10)
        frm.pack(fill=X)
        tb.Label(frm, text="Sources à sauvegarder:", font=(None,12,'bold')).grid(row=0,column=0,sticky=W)
        tb.Button(frm, text="Ajouter", bootstyle=SUCCESS, command=self.ajouter_source).grid(row=0,column=1,**pad)
        tb.Button(frm, text="Retirer", bootstyle=DANGER, command=self.retirer_source).grid(row=0,column=2)
        self.lst_sources = tk.Listbox(frm, height=5)
        self.lst_sources.grid(row=1,column=0,columnspan=3,sticky=EW,**pad)

        frm2 = tb.Frame(self, padding=10)
        frm2.pack(fill=X)
        tb.Label(frm2, text="Destination NAS:", font=(None,12,'bold')).grid(row=0,column=0,sticky=W)
        self.ent_dst = tb.Entry(frm2, width=50)
        self.ent_dst.grid(row=0,column=1,**pad)
        tb.Button(frm2, text="Parcourir", command=self.choisir_destination).grid(row=0,column=2)

        opts = tb.Labelframe(self, text="Options NAS & Système", padding=10)
        opts.pack(fill=X,**pad)
        tb.Label(opts, text="Utilisateur:").grid(row=0,column=0,sticky=W)
        self.ent_user = tb.Entry(opts, width=20); self.ent_user.insert(0,"admin"); self.ent_user.grid(row=0,column=1,**pad)
        tb.Label(opts, text="Mot de passe:").grid(row=1,column=0,sticky=W)
        self.ent_pwd = tb.Entry(opts, show="*", width=20); self.ent_pwd.grid(row=1,column=1)
        self.chk_sys = tb.Checkbutton(opts, text="Inclure image système", bootstyle=INFO); self.chk_sys.grid(row=0,column=2,rowspan=2,padx=20)
        tb.Label(opts, text="Partition:").grid(row=2,column=0,sticky=W)
        self.ent_part = tb.Entry(opts, width=10); self.ent_part.insert(0,"C:"); self.ent_part.grid(row=2,column=1)

        sch = tb.Labelframe(self, text="Planification", padding=10)
        sch.pack(fill=X,**pad)
        tb.Label(sch, text="Heure HH:MM:").grid(row=0,column=0,sticky=W)
        self.var_heure = tb.StringVar(); tb.Entry(sch, textvariable=self.var_heure, width=10).grid(row=0,column=1,**pad)
        tb.Label(sch, text="Fréquence:").grid(row=1,column=0,sticky=W)
        self.var_freq = tb.StringVar(value="Jamais")
        tb.Combobox(sch, textvariable=self.var_freq, values=["Jamais","Tous les jours","Toutes les semaines"], state="readonly").grid(row=1,column=1)

        act = tb.Frame(self, padding=10); act.pack(fill=X)
        tb.Button(act, text="Lancer", bootstyle=PRIMARY, command=self.lancer_sauvegarde).pack(side=LEFT,**pad)
        tb.Button(act, text="Réinitialiser", bootstyle=WARNING, command=self.reinitialiser).pack(side=LEFT)

        self.pb = tb.Progressbar(self, length=600, bootstyle="info-striped")
        self.pb.pack(pady=20)
        self.lbl_status = tb.Label(self, text="En attente...", font=(None,11))
        self.lbl_status.pack()

    # ----------------- Actions ----------------- #
    def ajouter_source(self):
        d = filedialog.askdirectory()
        if d:
            self.sources.append(d)
            self.lst_sources.insert(tk.END, d)

    def retirer_source(self):
        sel = self.lst_sources.curselection()
        if sel:
            idx = sel[0]
            self.sources.pop(idx)
            self.lst_sources.delete(idx)

    def choisir_destination(self):
        d = filedialog.askdirectory()
        if d:
            self.ent_dst.delete(0, tk.END)
            self.ent_dst.insert(0, d)

    def lancer_sauvegarde(self):
        if not self.sources:
            messagebox.showerror("Erreur","Aucune source définie.")
            return
        dst = self.ent_dst.get().strip()
        if not dst or not os.path.isdir(dst):
            messagebox.showerror("Erreur","Destination invalide.")
            return
        threading.Thread(target=self.sauvegarde_thread, daemon=True).start()

    def sauvegarde_thread(self):
        self.log("Démarrage sauvegarde…")
        erreurs = []
        dst_root = self.ent_dst.get().strip()
        self.connecter_nas(dst_root, self.ent_user.get(), self.ent_pwd.get())
        total_count = [self.compter_fichiers()]
        counter = [0]

        if self.chk_sys.instate(['selected']):
            self.lbl_status.config(text="Image système…")
            self.pb.start(10)
            self.sauvegarder_image_systeme(dst_root, self.ent_part.get())
            self.pb.stop()
            self.lbl_status.config(text="Image système OK")

        for src in self.sources:
            if not os.path.isdir(src):
                erreurs.append(src)
                continue
            nom = os.path.basename(src.rstrip(os.sep))
            last = os.path.join(dst_root, f"{nom}_inc_last")
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
            snapshot = os.path.join(dst_root, f"{nom}_inc_{timestamp}")

            if not self.verifier_destination_valide(src, snapshot):
                return

            self.copier_incrementiel(src, last, total_count, counter)

            if os.path.exists(snapshot):
                shutil.rmtree(snapshot)
            os.rename(last, snapshot)
            self.creer_snapshot_hardlink(snapshot, last)

        if erreurs:
            messagebox.showwarning("Partielle", "Erreurs:" + "\n".join(erreurs))
            self.lbl_status.config(text="Terminé avec erreurs ⚠️")
        else:
            messagebox.showinfo("Succès","Sauvegarde terminée ✅")
            self.lbl_status.config(text="Terminé ✅")

    def reinitialiser(self):
        self.sources.clear()
        self.lst_sources.delete(0,tk.END)
        self.ent_dst.delete(0,tk.END)
        self.ent_user.delete(0,tk.END)
        self.ent_pwd.delete(0,tk.END)
        self.ent_part.delete(0,tk.END)
        self.var_heure.set("")
        self.var_freq.set("Jamais")
        self.chk_sys.deselect()
        self.pb.stop()
        self.pb['value'] = 0
        self.lbl_status.config(text="En attente...")

    def planification_automatique(self):
        while True:
            freq, heure = self.var_freq.get(), self.var_heure.get()
            if freq != 'Jamais' and heure:
                try:
                    h,m = map(int, heure.split(':'))
                    now = datetime.now()
                    cible = now.replace(hour=h, minute=m, second=0, microsecond=0)
                    if now >= cible:
                        ajout = 1 if freq == 'Tous les jours' else 7
                        cible += timedelta(days=ajout)
                    threading.Event().wait((cible-now).total_seconds())
                    self.lancer_sauvegarde()
                except Exception as e:
                    self.log(f"Erreur planification: {e}")
            else:
                threading.Event().wait(60)

if __name__ == '__main__':
    app = BackupApp()
    app.mainloop()
