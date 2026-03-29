# Microsoft Store Submission & MSIX Testing Guide

This guide describes how to use the newly created MSIX packaging system for **ToHost Bridge** to submit your app to the Microsoft Store and test it on your machine.

## 1. Local Testing
I have created a **self-signed certificate** (`msix/ToHostBridgeDev.pfx`) so you can test the installation process immediately.

### Install the App
-   Run `make_msix.ps1` on your build machine to generate the `.msix` in `releases/`.
-   Double-click `ToHostBridge_v1.1.0_x64.msix` on the target machine.

## 2. Microsoft Store Submission (Partner Center)
To release on the Store, follow these steps:

### A. Reserve App Name
1.  Log in to the [Microsoft Partner Center](https://partner.microsoft.com/dashboard/desktop).
2.  Select **Apps and games** > **New product**.
3.  Enter **ToHost Bridge** and reserve the name.

### B. Update Manifest Identity
1.  Go to **Product management** > **Product Identity**.
2.  Copy these values into your `msix/Package.appxmanifest`:
    -   **Package/Identity/Name** (e.g., `MaxCoppola.ToHostBridge`)
    -   **Package/Identity/Publisher** (e.g., `CN=B2012D8E-C27E-4748-825E-1B3C39DEB0BF`)
3.  **IMPORTANT**: The `Publisher` string in the manifest **MUST** exactly match the one in Partner Center for the Store to accept the package.

### C. Build for Store
1.  Run `make_msix.ps1` to re-generate the package with the correct identity.
2.  **Note**: The Store will re-sign the package for you, so the local self-signed cert is only for testing.

### D. Upload and Submit
1.  In Partner Center, create a **New Submission**.
2.  Upload your `.msix` from the `releases/` folder.
3.  Fill in the Store Listing (Description, Categories, Brand Icons).
4.  **Submission Notes**: Mention that this is a "Desktop Bridge" app requiring Serial and MIDI access for legacy hardware connectivity.

## Tool Locations
-   **Icons**: [msix/Assets/](file:///c:/Users/admin/.gemini/antigravity/scratch/ToHost%20Bridge%20V1.0/msix/Assets/) contains the generated premium store-ready icons.
-   **Manifest**: [msix/Package.appxmanifest](file:///c:/Users/admin/.gemini/antigravity/scratch/ToHost%20Bridge%20V1.0/msix/Package.appxmanifest).
-   **Automated Script**: [make_msix.ps1](file:///c:/Users/admin/.gemini/antigravity/scratch/ToHost%20Bridge%20V1.0/make_msix.ps1).
