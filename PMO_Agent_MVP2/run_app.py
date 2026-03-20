import asyncio
import sys
import streamlit.web.cli as stcli
from dotenv import load_dotenv

if __name__ == '__main__':
    if sys.platform == 'win32':
        try:
            asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
        except AttributeError:
            pass

    # Fix for Python 3.10+ where get_event_loop() doesn't auto-create loops
    try:
        asyncio.get_event_loop()
    except RuntimeError:
        asyncio.set_event_loop(asyncio.new_event_loop())

    load_dotenv()
    sys.argv = ["streamlit", "run", "app.py"]
    sys.exit(stcli.main())
