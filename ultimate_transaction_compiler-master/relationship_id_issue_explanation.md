# Understanding the Relationship ID Issue

## The Problem

Our transaction compiler had an issue where the same donor might have different relationship IDs in different rows of our output data. Specifically:

- One row might show a single ID: `10933483`
- Another row might show a concatenated ID: `10938370 + 10933483`

This inconsistency was problematic because these IDs should match across all rows for the same relationship.

## How Relationship IDs Work

To understand the issue, let's first understand how relationship IDs are supposed to work:

1. When we process donation data from different platforms (like EveryAction and ActBlue), we need to identify when the same donor appears in multiple systems.

2. We use "keys" (like email addresses or unique identifiers) to match records across platforms.

3. When we find a match, we assign a "relationship ID" to connect these records.

4. If a donor has multiple IDs across systems, we combine them with a "+" symbol (e.g., `10938370 + 10933483`).

## What Went Wrong

The issue was in the sequence of operations. Imagine we're processing a stack of donation records one by one:

1. We pick up the first record for donor Jane (from EveryAction) and assign ID `10933483`.

2. We continue processing other records.

3. Later, we find another record for Jane (from ActBlue) with ID `10938370`.

4. We realize these belong to the same person, so we create a combined ID: `10938370 + 10933483`.

5. We assign this combined ID to the ActBlue record.

**The problem**: We never went back to update the first EveryAction record with the combined ID. So Jane ended up with two different IDs in our system.

## The Solution: A Three-Phase Approach

The solution was to restructure our process into three distinct phases:

### Phase 1: Gather Information
- Process all records and collect information about each one
- Build lookup tables to help us find connections
- Don't assign any relationship IDs yet

### Phase 2: Make All Connections
- Go through all records again
- Identify all connections between records
- Build a complete map of which keys should have which relationship IDs
- Combine IDs where necessary
- Still don't assign IDs to the actual records

### Phase 3: Apply Final IDs
- Only after we have a complete picture of all connections
- Go through all records one last time
- Assign the final, fully combined relationship IDs to each record

## The Analogy

Think of it like organizing a family reunion:

**Old approach (problematic):**
1. As each person arrives, you give them a name tag
2. When you discover two people are related, you update their name tags to show their family connection
3. But if one person already left the registration desk, their name tag never gets updated

**New approach (solution):**
1. As people arrive, you just take their information but don't give name tags yet
2. You spend time figuring out all the family connections
3. Only after everyone has arrived and all connections are mapped out, you distribute name tags with the complete family information

## Benefits of the New Approach

1. **Consistency**: All related records now have the same relationship ID
2. **Accuracy**: We capture all connections properly
3. **Reliability**: The process works regardless of the order in which records are processed

This approach ensures that if Jane donated through multiple platforms, all her records will consistently show the same relationship ID, making it easier to track donor activity across systems.
