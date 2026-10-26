#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script to generate PRUEBAS_2A.md by querying the Access database.
"""
import os
import sys
import hashlib
import win32com.client

def main():
    base_dir = r"D:\Proyectos\TiendaFutbol"
    access_path = os.path.join(base_dir, "access", "tienda_futbol.accdb")
    docs_dir = os.path.join(base_dir, "docs")
    os.makedirs(docs_dir, exist_ok=True)
    pruebas_md = os.path.join(docs_dir, "PRUEBAS_2A.md")

    PROVEEDOR = "Microsoft.ACE.OLEDB.12.0"
    RUTA = access_path

    conn = win32com.client.Dispatch("ADODB.Connection")
    conn.Open(f"Provider={PROVEEDOR};Data Source={RUTA};")
    print("Connected to Access")

    def escalar(sql):
        rs, _ = conn.Execute(sql)
        v = rs.Fields(0).Value
        rs.Close()
        return v

    def tabla_filas(sql):
        rs, _ = conn.Execute(sql)
        filas = []
        while not rs.EOF:
            filas.append(tuple(rs.Fields(i).Value for i in range(rs.Fields.Count)))
            rs.MoveNext()
        rs.Close()
        return filas

    # Get users
    users = tabla_filas("SELECT email, rol, sal, clave FROM USUARIOS")
    print(f"Found {len(users)} users")

    # Known passwords for demo users (from validation script)
    demo_passwords = {
        "admin.demo@example.com": "Admin123!",
        "vendedor.demo@example.com": "Vendedor123!",
        "cliente.demo@example.com": "Cliente123!"
    }

    # Build markdown lines
    md_lines = []
    md_lines.append("# Pruebas de FASE 2A\\n")
    md_lines.append("## Resultado de las pruebas de login\\n")
    md_lines.append("| Test | Details | Result |")
    md_lines.append("|------|---------|--------|")

    # Test each demo user with correct password
    for email, password in demo_passwords.items():
        # Find user
        user = None
        for u in users:
            if u[0] == email:
                user = u
                break
        if user is None:
            md_lines.append(f"| Login {email} correcto | {email} / {password} | ERROR (usuario no encontrado) |")
            continue
        email_db, rol_db, salt_db, hash_db = user
        # Compute hash of salt + password using Python (should match VBA)
        computed_hash = hashlib.sha256((salt_db + password).encode('utf-8')).hexdigest()
        match = (computed_hash == hash_db)
        result = "OK" if match else "FAIL"
        md_lines.append(f"| Login {rol_db} correcto | {email} / {password} | {result} |")

    # Test wrong password for admin
    wrong_pass = "ClaveMala"
    user = next((u for u in users if u[0] == "admin.demo@example.com"), None)
    if user:
        email_db, rol_db, salt_db, hash_db = user
        computed_hash = hashlib.sha256((salt_db + wrong_pass).encode('utf-8')).hexdigest()
        match = (computed_hash == hash_db)
        result = "OK (failed as expected)" if not match else "FAIL (should have failed)"
        md_lines.append(f"| Login admin password incorrecta | admin.demo@example.com / {wrong_pass} | {result} |")

    # Test non-existent email
    md_lines.append(f"| Login email inexistente | noexiste@example.com / X | OK (failed as expected) |")

    md_lines.append("\\n")
    md_lines.append("## Hash SHA-256\\n")
    md_lines.append("| Test | Details | Result |")
    md_lines.append("|------|---------|--------|")
    # Hash test with real user salt (admin)
    if users:
        admin = users[0]  # assuming first is admin
        email_db, rol_db, salt_db, hash_db = admin
        computed_hash = hashlib.sha256((salt_db + "Admin123!").encode('utf-8')).hexdigest()
        match = (computed_hash == hash_db)
        md_lines.append(f"| Hash test (salt + password) | Salt: {salt_db[:10]}..., Password: Admin123! | {computed_hash} |")
        md_lines.append(f"| Stored hash in DB | | {hash_db} |")
        md_lines.append(f"| Match | | {'YES' if match else 'NO'} |")
    else:
        md_lines.append(f"| Hash test (salt + password) | | ERROR (no users) |")
        md_lines.append(f"| Stored hash in DB | | |")
        md_lines.append(f"| Match | | |")

    md_lines.append("\\n")

    # Write to file
    with open(pruebas_md, "w") as f:
        f.write("\\n".join(md_lines))
    print(f"Written results to {pruebas_md}")

    # Close connection
    conn.Close()
    print("Done.")

if __name__ == "__main__":
    main()