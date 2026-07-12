# Agent Task Planner Samples

These samples document the expected decomposition behavior for one chatbox message containing multiple requests.

## CSI Workflow

Input:

```text
Add frames from 0,0,0 to 6000,0,0, then assign section UC203x203x46, then extract frame info.
```

Expected task list:

1. Add frames from 0,0,0 to 6000,0,0
2. assign section UC203x203x46
3. extract frame info

## Mixed Query And Create

Input:

```text
Check selected points, create missing points, and summarize errors.
```

Expected task list:

1. Check selected points
2. create missing points
3. summarize errors

## Vietnamese Separators

Input:

```text
Tạo Howe truss 6 bays span 12000 rồi gán section W12X26 và sau đó list frame names.
```

Expected task list:

1. Tạo Howe truss 6 bays span 12000
2. gán section W12X26
3. list frame names

## Runtime Clarification

Input:

```text
Read this request, update the UI, and also fix the service naming.
```

Expected task list:

1. Read this request
2. update the UI
3. fix the service naming

In the chatbox runtime, task 2 and task 3 should be reported as requiring clarification or a development workflow, not as completed model/API changes.
