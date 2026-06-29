import pytest
from datetime import date, timedelta
from src.excel_reader import LeaveRecord
from src.email_sender import OutlookSender

def test_leave_duration_calculation():
    sender = OutlookSender()
    
    # Test a 1-day leave
    today = date.today()
    leave = LeaveRecord("Test User", "Planned", today, today, "Test")
    assert sender._duration(leave) == "1 day (today)"
    
    # Test a 3-day leave
    start = today
    end = today + timedelta(days=2)
    leave_long = LeaveRecord("Test User", "Planned", start, end, "Test")
    assert sender._duration(leave_long) == "3 day(s)"

def test_filter_today():
    sender = OutlookSender()
    today = date.today()
    
    # Record that starts today
    l1 = LeaveRecord("User1", "Planned", today, today + timedelta(days=1), "Reason")
    # Record that happened in the past
    l2 = LeaveRecord("User2", "Planned", today - timedelta(days=5), today - timedelta(days=2), "Reason")
    
    leaves = [l1, l2]
    filtered = sender._filter_today(leaves)
    
    assert len(filtered) == 1
    assert filtered[0].employee_name == "User1"