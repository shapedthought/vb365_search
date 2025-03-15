"""
Tests for the Exchange search functionality
"""

import unittest
from unittest.mock import patch, MagicMock

from vb365_search.search.exchange import ExchangeSearch
from vb365_search.search.models import ExchangeSearchResponse


class TestExchangeSearch(unittest.TestCase):
    """
    Test the Exchange search functionality
    """
    
    def setUp(self):
        """
        Set up the test case
        """
        self.config = {
            "vb365": {
                "api_address": "test-server",
                "version": "v8"
            }
        }
        self.auth_headers = {"Authorization": "Bearer test-token"}
        self.restore_session_id = "test-session-id"
        
        self.search = ExchangeSearch(
            self.config,
            self.auth_headers,
            self.restore_session_id
        )
    
    @patch("vb365_search.search.exchange.requests.post")
    def test_search(self, mock_post):
        """
        Test the search method
        """
        # Mock the response
        mock_response = MagicMock()
        mock_response.json.return_value = {
            "offset": 0,
            "limit": 30,
            "setId": "test-set-id",
            "results": [
                {
                    "mailboxId": "test-mailbox-id",
                    "attachments": [],
                    "organizer": "test-organizer",
                    "attendees": "test-attendees",
                    "startTime": "2023-01-01T00:00:00Z",
                    "endTime": "2023-01-01T01:00:00Z",
                    "location": "test-location",
                    "subject": "Test Subject",
                    "recurrencePatternFormat": "",
                    "recurring": False,
                    "itemClass": "test-item-class",
                    "_links": {
                        "property1": {"href": "test-href-1"},
                        "property2": {"href": "test-href-2"}
                    },
                    "_actions": {
                        "property1": {"uri": "test-uri-1", "method": "GET"},
                        "property2": {"uri": "test-uri-2", "method": "POST"}
                    },
                    "id": "test-id",
                    "name": "test-name",
                    "address": "test-address",
                    "businessPhone": "test-business-phone",
                    "company": "test-company",
                    "displayAs": "test-display-as",
                    "email": "test@example.com",
                    "fax": "test-fax",
                    "fileAs": "test-file-as",
                    "fullName": "Test User",
                    "homePhone": "test-home-phone",
                    "imAddress": "test-im-address",
                    "jobTitle": "test-job-title",
                    "mobile": "test-mobile",
                    "webPage": "test-web-page",
                    "from": "sender@example.com",
                    "postedOn": "2023-01-01T00:00:00Z",
                    "importance": "Normal",
                    "cc": "cc@example.com",
                    "bcc": "bcc@example.com",
                    "to": "recipient@example.com",
                    "sent": "2023-01-01T00:00:00Z",
                    "received": "2023-01-01T00:00:01Z",
                    "reminder": False,
                    "duration": 3600,
                    "entryType": "test-entry-type",
                    "date": "2023-01-01",
                    "status": "test-status",
                    "percentComplete": 0,
                    "startDate": "2023-01-01",
                    "dueDate": "2023-01-02",
                    "owner": "test-owner"
                }
            ],
            "_links": {
                "property1": {"href": "test-href-1"},
                "property2": {"href": "test-href-2"}
            }
        }
        mock_post.return_value = mock_response
        
        # Call the method
        result = self.search.search("test query")
        
        # Verify the result
        self.assertIsInstance(result, ExchangeSearchResponse)
        self.assertEqual(len(result.results), 1)
        self.assertEqual(result.results[0].subject, "Test Subject")
        self.assertEqual(result.results[0].from_, "sender@example.com")
        self.assertEqual(result.results[0].to, "recipient@example.com")
        
        # Verify the API call
        mock_post.assert_called_once()
        args, kwargs = mock_post.call_args
        self.assertEqual(
            args[0],
            "https://test-server:4443/v8/RestoreSessions/test-session-id/organization/mailboxes/search?limit=30"
        )
        self.assertEqual(kwargs["headers"], self.auth_headers)
        self.assertEqual(kwargs["verify"], False)
    
    def test_get_results(self):
        """
        Test the get_results method
        """
        # Set up test data
        self.search.results = {
            "results": [
                {
                    "subject": "Test Subject",
                    "received": "2023-01-01T00:00:01Z",
                    "from": "sender@example.com",
                    "to": "recipient@example.com",
                    "cc": "cc@example.com",
                    "bcc": "bcc@example.com",
                    "importance": "Normal",
                    "attachments": [
                        {"name": "test.txt", "sizeBytes": 1024}
                    ],
                    "id": "test-id",
                    # Add other required fields
                    "mailboxId": "test-mailbox-id",
                    "organizer": "test-organizer",
                    "attendees": "test-attendees",
                    "startTime": "2023-01-01T00:00:00Z",
                    "endTime": "2023-01-01T01:00:00Z",
                    "location": "test-location",
                    "recurrencePatternFormat": "",
                    "recurring": False,
                    "itemClass": "test-item-class",
                    "_links": {
                        "property1": {"href": "test-href-1"},
                        "property2": {"href": "test-href-2"}
                    },
                    "_actions": {
                        "property1": {"uri": "test-uri-1", "method": "GET"},
                        "property2": {"uri": "test-uri-2", "method": "POST"}
                    },
                    "name": "test-name",
                    "address": "test-address",
                    "businessPhone": "test-business-phone",
                    "company": "test-company",
                    "displayAs": "test-display-as",
                    "email": "test@example.com",
                    "fax": "test-fax",
                    "fileAs": "test-file-as",
                    "fullName": "Test User",
                    "homePhone": "test-home-phone",
                    "imAddress": "test-im-address",
                    "jobTitle": "test-job-title",
                    "mobile": "test-mobile",
                    "webPage": "test-web-page",
                    "postedOn": "2023-01-01T00:00:00Z",
                    "sent": "2023-01-01T00:00:00Z",
                    "reminder": False,
                    "duration": 3600,
                    "entryType": "test-entry-type",
                    "date": "2023-01-01",
                    "status": "test-status",
                    "percentComplete": 0,
                    "startDate": "2023-01-01",
                    "dueDate": "2023-01-02",
                    "owner": "test-owner"
                }
            ]
        }
        
        # Call the method
        results = self.search.get_results()
        
        # Verify the results
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0]["subject"], "Test Subject")
        self.assertEqual(results[0]["from"], "sender@example.com")
        self.assertEqual(results[0]["to"], "recipient@example.com")
        self.assertEqual(results[0]["has_attachments"], True)
        self.assertEqual(results[0]["attachments"], ["test.txt"])


if __name__ == "__main__":
    unittest.main()
