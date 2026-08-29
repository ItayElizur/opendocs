using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeAi.Shared
{
    public struct FreeSlot
    {
        public DateTime Start;
        public DateTime End;
        public int Available;   // attendees with no conflict
        public int Total;       // attendees checked
        public string[] Missing; // labels of attendees who ARE busy in this slot
    }

    // Pure, COM-free ranking of candidate meeting times from Outlook
    // Recipient.FreeBusy strings. Lives here (not in OutlookAiAddIn) so it can
    // be unit tested without the Outlook PIA. See OutlookTools.Calendar.cs's
    // find_meeting_slots for the COM half (resolving recipients, calling
    // Recipient.FreeBusy).
    //
    // Each FreeBusy string has one character per `stepMinutes` minutes,
    // starting at `rangeStartMidnight`. '0' = free; any other char (or an
    // index past the end of the string, or before the range) is treated as
    // free too only when past the end - a shorter-than-expected string means
    // "no known busy info". Before the range is never reached in practice.
    public static class MeetingSlots
    {
        public static List<FreeSlot> Rank(
            Dictionary<string, string> freeBusyByAttendee,
            DateTime rangeStartMidnight,
            IEnumerable<DateTime> businessDays,
            int startHour,
            int endHour,
            int durationMinutes,
            int stepMinutes,
            int limit)
        {
            if (stepMinutes <= 0) stepMinutes = 30;
            if (endHour <= startHour) throw new ArgumentException("end_hour must be after start_hour.");
            if (durationMinutes <= 0) throw new ArgumentException("duration_minutes must be positive.");

            int slotCount = (int)Math.Ceiling(durationMinutes / (double)stepMinutes);
            var duration = TimeSpan.FromMinutes(durationMinutes);
            var step = TimeSpan.FromMinutes(stepMinutes);
            int total = freeBusyByAttendee.Count;

            var candidates = new List<FreeSlot>();
            foreach (DateTime day in businessDays)
            {
                DateTime windowStart = day.Date.AddHours(startHour);
                DateTime windowEnd = day.Date.AddHours(endHour);
                for (DateTime slotStart = windowStart; slotStart + duration <= windowEnd; slotStart += step)
                {
                    var missing = new List<string>();
                    foreach (var kv in freeBusyByAttendee)
                    {
                        if (!IsFree(kv.Value, rangeStartMidnight, slotStart, slotCount, stepMinutes))
                            missing.Add(kv.Key);
                    }
                    candidates.Add(new FreeSlot
                    {
                        Start = slotStart,
                        End = slotStart + duration,
                        Available = total - missing.Count,
                        Total = total,
                        Missing = missing.ToArray(),
                    });
                }
            }

            return candidates
                .OrderByDescending(c => c.Available)
                .ThenBy(c => c.Start)
                .Take(Math.Max(1, limit))
                .ToList();
        }

        private static bool IsFree(string freeBusy, DateTime rangeStartMidnight, DateTime slotStart, int slotCount, int stepMinutes)
        {
            if (string.IsNullOrEmpty(freeBusy)) return true;
            int startIdx = (int)Math.Round((slotStart - rangeStartMidnight).TotalMinutes / stepMinutes);
            for (int j = startIdx; j < startIdx + slotCount; j++)
            {
                if (j < 0) return false;
                if (j >= freeBusy.Length) return true; // past the known window - assume free
                if (freeBusy[j] != '0') return false;
            }
            return true;
        }
    }
}
