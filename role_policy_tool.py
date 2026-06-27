# -*- coding: utf-8 -*-
"""Generate Ed25519 keys for signed ClinicManager role policies."""

import argparse
from pathlib import Path

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey


def generate_keypair(private_key_path, public_key_path):
    private_key = Ed25519PrivateKey.generate()
    public_key = private_key.public_key()

    Path(private_key_path).write_bytes(
        private_key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.PKCS8,
            encryption_algorithm=serialization.NoEncryption(),
        )
    )
    Path(public_key_path).write_bytes(
        public_key.public_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PublicFormat.SubjectPublicKeyInfo,
        )
    )


def main():
    parser = argparse.ArgumentParser(description="ClinicManager role policy key helper")
    parser.add_argument("--generate-keypair", action="store_true", help="Generate an Ed25519 keypair")
    parser.add_argument("--private-key", default="role_policy_private_key.pem")
    parser.add_argument("--public-key", default="role_policy_public_key.pem")
    args = parser.parse_args()

    if not args.generate_keypair:
        parser.error("Use --generate-keypair")

    generate_keypair(args.private_key, args.public_key)
    print(f"Private key written to: {args.private_key}")
    print(f"Public key written to: {args.public_key}")
    print("Keep the private key only on the trusted admin computer.")


if __name__ == "__main__":
    main()
