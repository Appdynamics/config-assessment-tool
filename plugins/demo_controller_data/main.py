import logging
from itertools import islice

def run_plugin(context):
    """
    Demo Hybrid plugin entry point
    """
    logging.info("Demo Hybrid Plugin: Started execution.")
    job_name = context.get('jobFileName')
    logging.info(f"Demo Hybrid Plugin: Processing job '{job_name}'")

    controller_data = context.get("controllerData")
    if controller_data:
        logging.info("Printing first 2 items of controllerData:")
        for key, value in islice(controller_data.items(), 2):
            logging.info(f"Key: {key}, Value: {value}")
    else:
        logging.warning("controllerData not found in context")

    # Simulate some work
    result = {
        "plugin_name": "DemoHybridPlugin",
        "processed_job": job_name,
        "status": "success"
    }

    logging.info(f"Demo Hybrid Plugin: Finished execution with result: {result}")
    return result

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    from collections import OrderedDict

    # Simple dummy execution for manual testing without argparse
    print("Running in Standalone Mode (Demo Hybrid Plugin)")

    # Mock data for standalone test
    mock_data = OrderedDict([
        ("item1", "value1"),
        ("item2", "value2"),
        ("item3", "value3")
    ])

    context = {
        "jobFileName": "Standalone_CLI_Run",
        "controllerData": mock_data
    }
    run_plugin(context)
