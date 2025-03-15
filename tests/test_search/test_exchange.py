"""
Tests for the Exchange search functionality
"""

import pytest

from vb365_search.search.exchange import ExchangeItemsInMailboxesSearch
from vb365_search.search.models import ExchangeItemsInMailboxesResponse


# Using mock_config from conftest.py


@pytest.fixture
def auth_headers():
    """
    Fixture for authentication headers
    """
    return {"Authorization": "Bearer test-token"}


@pytest.fixture
def restore_session_id():
    """
    Fixture for restore session ID
    """
    return "test-session-id"


@pytest.fixture
def exchange_search(mock_config, auth_headers, restore_session_id):
    """
    Fixture for ExchangeItemsInMailboxesSearch instance
    """
    return ExchangeItemsInMailboxesSearch(
        mock_config,
        auth_headers,
        restore_session_id
    )


@pytest.fixture
def mock_search_response():
    """
    Fixture for mock search response
    """
    return {
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


def test_search(exchange_search, mock_search_response, mocker):
    """
    Test the search method
    """
    # Mock the requests.post method
    mock_response = mocker.Mock()
    mock_response.json.return_value = mock_search_response
    
    mock_post = mocker.patch("requests.post", return_value=mock_response)
    
    # Call the method
    result = exchange_search.search("test query")
    
    # Verify the result
    assert isinstance(result, ExchangeItemsInMailboxesResponse)
    assert len(result.results) == 1
    assert result.results[0].subject == "Test Subject"
    assert result.results[0].from_ == "sender@example.com"
    assert result.results[0].to == "recipient@example.com"
    
    # Verify the API call
    mock_post.assert_called_once()
    args, kwargs = mock_post.call_args
    assert args[0] == "https://test-server:4443/v8/RestoreSessions/test-session-id/organization/mailboxes/search?limit=30"
    assert kwargs["headers"] == exchange_search.auth_headers
    assert kwargs["verify"] is False


def test_get_results(exchange_search):
    """
    Test the get_results method
    """
    # Set up test data
    exchange_search.results = {
        "offset": 0,
        "limit": 30,
        "setId": "test-set-id",
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
        ],
        "_links": {
            "property1": {"href": "test-href-1"},
            "property2": {"href": "test-href-2"}
        }
    }
    
    # Call the method
    results = exchange_search.get_results()
    
    # Verify the results
    assert len(results) == 1
    assert results[0]["subject"] == "Test Subject"
    assert results[0]["from"] == "sender@example.com"
    assert results[0]["to"] == "recipient@example.com"
    assert results[0]["has_attachments"] is True
    assert results[0]["attachments"] == ["test.txt"]
