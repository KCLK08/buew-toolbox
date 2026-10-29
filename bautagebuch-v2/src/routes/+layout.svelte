<script>
  import '@fontsource/space-grotesk/400.css';
  import '@fontsource/space-grotesk/600.css';
  import '@fontsource/space-grotesk/700.css';
</script>

<svelte:head>
  <title>Bautagebuch</title>
  <meta
    name="description"
    content="Offline Bautagebuch mit PDF-Vorlage, Formularerkennung und Fotodokumentation"
  />
</svelte:head>

<style>
  :global(:root) {
    --bg: #f2f0eb;
    --ink: #1a1916;
    --muted: #6b6560;
    --accent: #c44b32;
    --accent-pressed: #a33d29;
    --accent-2: #243240;
    --panel: #fffcf7;
    --panel-elevated: #ffffff;
    --panel-soft: #fffcf7;
    --border: #e4ddd2;
    --border-strong: #d0c7b8;
    --success: #2f6b45;
    --warning: #9a6b12;
    --danger: #a12c24;
    --info: #2a5f8f;
    --badge-bg: rgba(196, 75, 50, 0.12);
    --overlay: rgba(26, 25, 22, 0.45);
    --radius: 16px;
    --radius-button: 14px;
    --radius-input: 14px;
    --shadow: 0 4px 10px rgba(26, 25, 22, 0.08);
    --touch: 44px;
    font-family: 'Space Grotesk', 'Helvetica Neue', Arial, sans-serif;
  }

  :global(body) {
    margin: 0;
    background: var(--bg);
    color: var(--ink);
  }

  :global(*) {
    box-sizing: border-box;
  }
</style>

<slot />
