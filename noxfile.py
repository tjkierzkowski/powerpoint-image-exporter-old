import nox

locations = "pptx_export", "cli", "tests", "noxfile.py", "pptx_image_exporter.py"
nox.options.sessions = "lint", "tests"


@nox.session(python=["3.8", "3.7", "3.6"])
def tests(session):
    session.run("poetry", "install", external=True)
    session.run("pytest", "--cov")


@nox.session(python=["3.8", "3.7"])
def lint(session):
    args = session.posargs or locations
    session.install("flake8", "flake8-black", "flake8-import-order")
    session.run("flake8", *args)


@nox.session(python="3.8")
def black(session):
    args = session.posargs or locations
    session.install("black")
    session.run("black", *args)
