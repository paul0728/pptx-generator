"""Allow running as: python -m pptx_generator"""
import sys
from .generator import main

sys.exit(main())
