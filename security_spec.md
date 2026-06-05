# Security Specification: SIGAP MTsN 2 Bombana Firestore Security Rules (Zero-Trust)

This specification defines the security invariants and testing scenarios for the Firestore database schema of SIGAP MTsN 2 Bombana.

## 1. Data Invariants

1. **Teacher Credentials Privacy**: Raw credentials (such as passwords) must NEVER be publicly readable. Password records must be isolated inside `/teachers/{nip}/private/password` and blocked from direct read operations.
2. **Role-Based Authorization**: Data write operations (for administrative resources like `students`, `teachers`, `classrooms`, `settings`) must match an active session of role `Admin`.
3. **Attendance Isolation**: Students can ONLY read their own individual attendance records (`attendance/{id}`). They must never read other students' attendance records or teacher attendances.
4. **Session Integrity**: Session documents in `/activeSessions/{authUid}` must map exactly to the authenticated user's `uid` and can only be created if they supply matching, valid credentials.

---

## 2. The "Dirty Dozen" Adversarial Payloads

The following payloads and requests represent unauthorized attempts to bypass security. Our rules must reject all of them with `PERMISSION_DENIED`.

### Pillar 1: Identity Spoofing & Credential Harvesting
1. **Unauthenticated Read on teacher credentials**:
   - **Path**: `GET /teachers/ADMIN001/private/password`
   - **Expected**: `PERMISSION_DENIED`
2. **Harvesting all teacher passwords**:
   - **Path**: `GET /teachers` (list queries from guests)
   - **Expected**: `PERMISSION_DENIED`
3. **Impersonating a student session mapping without NISN validation**:
   - **Path**: `CREATE /activeSessions/unauthorized_uid` with payload:
     ```json
     {
       "uid": "unauthorized_uid",
       "nisn": "999999999",
       "role": "Siswa"
     }
     ```
     (Note: `999999999` does not exist in `students`).
   - **Expected**: `PERMISSION_DENIED`

### Pillar 2: Privilege Escalation
4. **Adversarial write to active sessions with self-assigned admin privileges**:
   - **Path**: `CREATE /activeSessions/attacker_uid` with payload:
     ```json
     {
       "uid": "attacker_uid",
       "role": "Admin"
     }
     ```
   - **Expected**: `PERMISSION_DENIED`
5. **Directly writing teacher configuration with self-promoted role**:
   - **Path**: `CREATE /teachers/ATTACKER01` with payload:
     ```json
     {
       "nip": "ATTACKER01",
       "nama": "Attacker",
       "role": "Admin"
     }
     ```
   - **Expected**: `PERMISSION_DENIED`

### Pillar 3: Resource Poisoning & Denial of Wallet
6. **Injecting oversized keys or malicious document IDs**:
   - **Path**: `CREATE /appConfig/very_long_junk_character_id_exceeding_128_bytes_..."`
   - **Expected**: `PERMISSION_DENIED`
7. **Writing broken student record violating schema rules**:
   - **Path**: `CREATE /students/123456` with payload:
     ```json
     {
       "nisn": "",
       "nama": 1002,
       "kelas": "VIII"
     }
     ```
   - **Expected**: `PERMISSION_DENIED`

### Pillar 4: Lateral Access
8. **Student attempting to read attendance of another student**:
   - **Path**: `GET /attendance/studentB-2026-06-05` (when logged in as student A)
   - **Expected**: `PERMISSION_DENIED`
9. **Student attempting to write/edit attendance records**:
   - **Path**: `UPDATE /attendance/studentA-2026-06-05` with payload:
     ```json
     {
       "id": "studentA-2026-06-05",
       "status": "Hadir"
     }
     ```
   - **Expected**: `PERMISSION_DENIED`

### Pillar 5: Data Defacement & Orphaned Records
10. **Unauthenticated deletion of settings configuration**:
    - **Path**: `DELETE /settings/Senin`
    - **Expected**: `PERMISSION_DENIED`
11. **Altering school holidays without Admin credentials**:
    - **Path**: `CREATE /holidays/2026-08-17` with payload:
      ```json
      {
        "tanggal": "2026-08-17",
        "keterangan": "Defaced Holiday"
      }
      ```
    - **Expected**: `PERMISSION_DENIED`
12. **Clearing teacher attendance rekap files**:
    - **Path**: `DELETE /teacherAttendance/ADMIN001-2026-06-05`
    - **Expected**: `PERMISSION_DENIED`

---

## 3. Test Assertion Engine Mapping

Our security test wrapper validates that regardless of client-side queries, Firestore enforces authorization at the database level.

| Collection Path | Unauthenticated | Student Session | Teacher/Admin Session |
|---|---|---|---|
| `/activeSessions/{uid}` | Create Only (if qualified) | Read (Self Only) | Read/Write (Self Only) |
| `/teachers/{nip}` | Denied | Denied | Approved |
| `/teachers/{nip}/private/password` | Denied | Denied | Denied (Read/Write) |
| `/students/{nisn}` | Denied | Read (Self Only) | Approved |
| `/attendance/{id}` | Denied | Read (Self Only) | Approved |
| `/classrooms/{nama}` | Denied | Read | Approved |
| `/appConfig/{id}` | Read Only | Read Only | Approved |
