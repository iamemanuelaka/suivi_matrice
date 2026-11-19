"""Script d'entrée de l'exécutable Windows (PyInstaller) : lance l'application PCR-CI."""
import multiprocessing

if __name__ == "__main__":
    multiprocessing.freeze_support()
    from pcr_app.lanceur import main
    main()
