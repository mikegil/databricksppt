#!/usr/bin/env python

"""Tests for `tcppt` package."""


import unittest
from click.testing import CliRunner

from tcppt import tcppt
from tcppt import cli


class TestTcppt(unittest.TestCase):
    """Tests for `tcppt` package."""

    def setUp(self):
        """Set up test fixtures, if any."""

    def tearDown(self):
        """Tear down test fixtures, if any."""

    def test_000_something(self):
        print(tcppt.toTCPPT(""))

    def test_command_line_interface(self):
        """Test the CLI."""
        runner = CliRunner()
        result = runner.invoke(cli.main)
        assert result.exit_code == 0
        assert 'tcppt.cli.main' in result.output
        help_result = runner.invoke(cli.main, ['--help'])
        assert help_result.exit_code == 0
        assert '--help  Show this message and exit.' in help_result.output
