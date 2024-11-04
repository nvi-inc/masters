import os
import sys
import toml
import argparse


def init(args: argparse.Namespace) -> argparse.Namespace:
    """
    Get application input options and parameters
    :param args: argument provided by argparse init
    :return: same as input
    """
    # Initialize global variables using setattr
    this_module = sys.modules[__name__]
    setattr(this_module, 'args', args)  # Set default options for this app
    try:
        # Read config file in toml format
        setattr(this_module, 'config', toml.load(os.path.expanduser(args.config)))
        return args
    except FileNotFoundError as exc:
        print(f'Problem opening {args.config}! File not found')
    except Exception as exc:
        print(f'Problem reading {args.config} {str(exc)}')
    sys.exit(0)
