---
layout: default
title: Project Management
position: Project / Programme Mgt
---

## Project Management
Despite what leading PM software vendors would have you believe, there are short-comings in most of these tools. Persistent annoyances should be dealt with efficiently in case they affect productivity and deadlines:

- people take time-off, global projects are affected by different public holiday schedules and team members have training days and other legitimate absence, so efficiently dealing with change in project capacity is important. Most calendars are published as event streams, so understanding how to process these is a helpful skill ics-event-stream.vb

{% assign samplecode_files = site.static_files | where: "samplecode", true %}
{% for mycode in samplecode_files %}
  {% link mycode.path %}
{% endfor %}
