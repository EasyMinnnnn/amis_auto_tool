"""Streamlit application for automating AMIS downloads and document creation."""

from __future__ import annotations

import os
import tempfile
from datetime import datetime

import streamlit as st

# Import the helper functions from the amis module within this package.
from . import amis


def main() -> None:
    st.set_page_config(page_title="AMIS Automation", layout="centered")
    st.title("AMIS Auto‑Downloader")
    st.write(
        """
        This tool automates the process of downloading the property information
        sheet and associated photos from the AMIS portal.  Provide your AMIS
        credentials and the record ID you want to process, then upload your
        signature image.  The application will log in, retrieve the necessary
        files and assemble them into a single Word document for you to
        download.
        """
    )

    # User inputs
    username = st.text_input("AMIS username")
    password = st.text_input("AMIS password", type="password")
    record_id = st.text_input("Record ID")
    signature_file = st.file_uploader("Signature image (PNG/JPG)", type=["png", "jpg", "jpeg"])

    if st.button("Run automation"):
        # Validate inputs
        if not (username and password and record_id and signature_file):
            st.error("Please fill in all fields and upload a signature image.")
            return

        # Create a temporary working directory
        with tempfile.TemporaryDirectory() as tmpdir:
            downloads_dir = os.path.join(tmpdir, "downloads")
            os.makedirs(downloads_dir, exist_ok=True)

            # Save the uploaded signature to a temporary file
            sig_path = os.path.join(tmpdir, "signature.png")
            sig_bytes = signature_file.read()
            with open(sig_path, "wb") as f:
                f.write(sig_bytes)

            st.info("Connecting to AMIS and downloading files… this may take a few minutes.")
            try:
                template_path, images = amis.run_automation(
                    username=username,
                    password=password,
                    record_id=record_id,
                    download_dir=downloads_dir,
                    headless=True,
                )
            except Exception as e:
                st.exception(e)
                return

            st.success("Downloaded Word template and images. Preparing your document…")

            # Determine output path
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(tmpdir, f"Phieu_TTTT_{record_id}_{timestamp}.docx")

            try:
                amis.fill_document(
                    template_path=template_path,
                    images=images,
                    signature_path=sig_path,
                    output_path=output_path,
                )
            except Exception as e:
                st.exception(e)
                return

            # Read the final docx for download
            with open(output_path, "rb") as f:
                final_docx = f.read()

            st.success("Document ready!")
            st.download_button(
                label="Download completed Word document",
                data=final_docx,
                file_name=os.path.basename(output_path),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )


if __name__ == "__main__":
    main()
