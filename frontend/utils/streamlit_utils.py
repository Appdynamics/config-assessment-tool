import streamlit as st


def rerun():
    # no native way to do this at time of writing, so let's use a dirty hack
    # https://github.com/streamlit/streamlit/issues/653
    # raise st.script_runner.RerunException(st.script_request_queue.RerunData(None))
    st.experimental_rerun()
