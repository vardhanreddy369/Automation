---
description: Get Shit Done (GSD) Development Framework
---
# Get Shit Done (GSD) Development Framework

This workflow activates a highly-structured, spec-driven development loop. 

**Steps for the Agent:**
1. **Understand Requirements:** I will immediately look for and read `TODO.md` or `spec.md` in the root directory. If they don't exist, I will ask you to provide them. I will strictly refuse to write any code until I fully understand the requirements.
2. **Context Engineering & Planning:** I will formulate a step-by-step plan based on the specs and present it to you for approval.
3. **Sequential Execution:** Once approved, I will build out the plan one step at a time. I will execute and validate each step sequentially (using terminal commands, build tools, or test scripts) to prevent "context rot" and ensure everything works perfectly before moving on to the next step.
