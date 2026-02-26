import logging
import asyncio
import os

async def run_plugin(context):
    """
    Demo Hybrid plugin entry point
    """
    logging.info("Demo integrated plugin: Started execution.")
    job_name = context.get('jobFileName')
    logging.info(f"Processing job '{job_name}'")

    output_dir = context.get("outputDir")
    if output_dir:
        logging.info(f"Output directory found: {output_dir}")
        if os.path.exists(output_dir):
            try:
                files = os.listdir(output_dir)
                logging.info(f"Files found in output directory ({len(files)}):")
                for f in files:
                    logging.info(f" - {f}")
            except Exception as e:
                logging.error(f"Failed to list files in output directory: {e}")
        else:
            logging.warning(f"Output directory does not exist: {output_dir}")
    else:
        logging.warning("outputDir not found in context")

    logging.info(f"Demo Plugin: Finished execution")