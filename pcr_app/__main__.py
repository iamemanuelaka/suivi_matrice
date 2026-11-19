"""`python -m pcr_app` : lance l'application PCR-CI."""
from pcr_app.lanceur import main

if __name__ in {"__main__", "__mp_main__"}:
    main()
