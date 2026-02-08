"""
CLI entry point for the SHM Slide Layout Engine.

Usage:
    python run.py generate --manifest manifests/sample_security_report.csv
    python run.py generate --manifest manifests/report.csv --template templates/shm.pptx
    python run.py serve                       # start Flask API on port 5000
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from core.report_generator import ReportGenerator


def cmd_generate(args):
    gen = ReportGenerator(
        template_path=args.template,
        output_path=args.output,
        assets_dir=args.assets,
    )
    result = gen.generate(args.manifest)

    ok = sum(1 for r in gen.audit_trail if r.status == "ok")
    errors = sum(1 for r in gen.audit_trail if r.status == "error")
    print(f"Done. {ok} placed, {errors} errors. Output: {result}")


def cmd_serve(args):
    from api import app
    import os
    
    if args.production:
        # Production mode: use Waitress (cross-platform) or Gunicorn (Unix)
        print("🚀 Starting production server...")
        
        ssl_cert = args.ssl_cert or os.getenv("SSL_CERT_PATH")
        ssl_key = args.ssl_key or os.getenv("SSL_KEY_PATH")
        
        if ssl_cert and ssl_key:
            # HTTPS with proper SSL certificates
            print(f"🔒 HTTPS enabled with certificates:")
            print(f"   Cert: {ssl_cert}")
            print(f"   Key:  {ssl_key}")
            
            if os.name == 'nt':  # Windows
                # Use Waitress on Windows
                from waitress import serve
                serve(
                    app,
                    host="0.0.0.0",
                    port=args.port,
                    url_scheme='https',
                    ident=None  # Don't expose server version
                )
            else:  # Unix/Linux/Mac
                # Use Gunicorn on Unix (better performance)
                import subprocess
                cmd = [
                    "gunicorn",
                    "--bind", f"0.0.0.0:{args.port}",
                    "--workers", str(args.workers),
                    "--certfile", ssl_cert,
                    "--keyfile", ssl_key,
                    "--access-logfile", "-",
                    "--error-logfile", "-",
                    "api:app"
                ]
                print(f"   Workers: {args.workers}")
                print(f"   Running: {' '.join(cmd)}")
                subprocess.run(cmd)
        else:
            # HTTP production server (use reverse proxy like nginx for SSL)
            print("⚠️  Running without SSL - use nginx/Apache reverse proxy for HTTPS in production")
            
            if os.name == 'nt':  # Windows
                from waitress import serve
                serve(app, host="0.0.0.0", port=args.port, ident=None)
            else:  # Unix/Linux/Mac
                import subprocess
                cmd = [
                    "gunicorn",
                    "--bind", f"0.0.0.0:{args.port}",
                    "--workers", str(args.workers),
                    "--access-logfile", "-",
                    "--error-logfile", "-",
                    "api:app"
                ]
                print(f"   Workers: {args.workers}")
                subprocess.run(cmd)
    else:
        # Development mode: use Flask dev server
        ssl_context = None
        if args.ssl:
            try:
                # Use adhoc SSL (auto-generates a self-signed certificate)
                ssl_context = 'adhoc'
                print(f"🔒 Starting server with HTTPS on https://0.0.0.0:{args.port}")
                print("⚠️  Using self-signed certificate - your browser will show a warning (this is normal for development)")
            except ImportError:
                print("⚠️  pyOpenSSL not installed. Install with: pip install pyopenssl")
                print("    Falling back to HTTP...")
                ssl_context = None
        
        app.run(host="0.0.0.0", port=args.port, debug=args.debug, ssl_context=ssl_context)


def main():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(name)-24s  %(levelname)-8s  %(message)s",
    )

    parser = argparse.ArgumentParser(description="SHM Slide Layout Engine")
    sub = parser.add_subparsers(dest="command")

    # --- generate ---
    gen_p = sub.add_parser("generate", help="Generate a report from a manifest")
    gen_p.add_argument("--manifest", required=True, help="Path to layout manifest (CSV/JSON)")
    gen_p.add_argument("--template", default=None, help="Path to PPTX template")
    gen_p.add_argument("--output", default="output/report.pptx", help="Output PPTX path")
    gen_p.add_argument("--assets", default="assets", help="Base directory for images")

    # --- serve ---
    srv_p = sub.add_parser("serve", help="Start the Flask API server")
    srv_p.add_argument("--port", type=int, default=5000, help="Port to run server on")
    srv_p.add_argument("--debug", action="store_true", help="Enable debug mode (development only)")
    srv_p.add_argument("--ssl", action="store_true", help="Enable HTTPS with self-signed certificate (dev only)")
    srv_p.add_argument("--production", action="store_true", help="Run in production mode with Gunicorn/Waitress")
    srv_p.add_argument("--ssl-cert", help="Path to SSL certificate file (production)")
    srv_p.add_argument("--ssl-key", help="Path to SSL key file (production)")
    srv_p.add_argument("--workers", type=int, default=4, help="Number of worker processes (production, default: 4)")

    args = parser.parse_args()
    if args.command == "generate":
        cmd_generate(args)
    elif args.command == "serve":
        cmd_serve(args)
    else:
        parser.print_help()
        sys.exit(1)


if __name__ == "__main__":
    main()
