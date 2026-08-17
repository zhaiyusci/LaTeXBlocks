const repositoryUrl = "https://github.com/zhaiyusci/LaTeXBlocks";
const downloadUrl = `${repositoryUrl}/releases/latest`;

export default function Home() {
  return (
    <main>
      <nav className="nav shell" aria-label="Primary navigation">
        <a className="brand" href="#top" aria-label="LaTeX Blocks home">
          <img src="/brand-icon.png" alt="" width="42" height="42" />
          <span>LaTeX Blocks</span>
        </a>
        <div className="navLinks">
          <a href="#features">Features</a>
          <a href={`${repositoryUrl}#documentation`}>Docs</a>
          <a className="navCta" href={downloadUrl}>Download</a>
        </div>
      </nav>

      <section className="hero shell" id="top">
        <div className="heroCopy">
          <div className="eyebrow"><i /> Built for Word + PowerPoint</div>
          <h1>LaTeX that belongs<br />in your document.</h1>
          <p className="lede">
            Create precise, editable LaTeX blocks in Microsoft Office.
            Beautiful SVG on the surface. Your source stays with the object.
          </p>
          <div className="actions">
            <a className="button primary" href={downloadUrl}>Download for Windows <span>↓</span></a>
            <a className="button secondary" href={repositoryUrl}>View on GitHub <span>↗</span></a>
          </div>
          <div className="requirements">
            <span>Windows 64-bit</span><b>·</b><span>Office 64-bit</span><b>·</b><span>Open source</span>
          </div>
        </div>

        <div className="productStage" aria-label="LaTeX Blocks in an Office document">
          <div className="stageGlow" />
          <div className="officeCard">
            <div className="officeChrome">
              <div className="windowDots"><i /><i /><i /></div>
              <span>Research notes.docx</span>
              <div className="wordBadge">W</div>
            </div>
            <div className="ribbon">
              <span className="ribbonTab">LaTeX Blocks</span>
              <div className="ribbonTools"><i>∫</i><i>∑</i><i>▱</i><i>↗</i></div>
            </div>
            <div className="document">
              <p>For a quantum state with Hamiltonian</p>
              <div className="equationBox">
                <span>Ĥ</span> = <span className="fraction"><b>p̂²</b><i>2m</i></span> + V(q̂)
                <div className="handle tl" /><div className="handle tr" />
                <div className="handle bl" /><div className="handle br" />
              </div>
              <p>the partition function can be written as an editable block.</p>
              <div className="sourceTag"><span>Source</span><code>{"\\hat H = \\frac{\\hat p^2}{2m} + V(\\hat q)"}</code></div>
            </div>
          </div>
          <div className="formatChip"><i>✓</i><span><b>Portable SVG</b><small>Looks right everywhere</small></span></div>
        </div>
      </section>

      <section className="trustStrip" id="features">
        <div className="shell trustInner">
          <p><strong>Real TeX layout.</strong> Native Office workflow.</p>
          <div><span>WORD</span><i>+</i><span>POWERPOINT</span><i>+</i><span>STEMTEX</span></div>
        </div>
      </section>

      <section className="section shell capabilities">
        <div className="sectionIntro">
          <span className="kicker">Made for serious documents</span>
          <h2>Office-native on the outside.<br />LaTeX all the way through.</h2>
          <p>Every block carries its own source, so editing stays effortless and the document remains portable.</p>
        </div>
        <div className="featureGrid">
          <article className="feature featureWide navyFeature">
            <div className="featureNumber">01</div>
            <div>
              <h3>Edit the source, not the pixels.</h3>
              <p>Double-click a recognized object, change the author LaTeX, and render it back into the same Office frame.</p>
            </div>
            <div className="editorMini">
              <div><span>LaTeX source</span><b>Preview is current</b></div>
              <code>{"\\begin{align}\n  Z &= \\int e^{-\\beta H} \\, dq \\, dp\n\\end{align}"}</code>
              <div className="previewEquation">Z = ∫ e<sup>−βH</sup> dq dp</div>
            </div>
          </article>
          <article className="feature">
            <div className="featureIcon sourceIcon">{`{ }`}</div>
            <h3>Source travels with it</h3>
            <p>LaTeX is persisted on the visual object—not in a hidden companion file.</p>
          </article>
          <article className="feature">
            <div className="featureIcon resizeIcon">↔</div>
            <h3>Resize, then reflow</h3>
            <p>Release the frame and the block is typeset again for its exact new dimensions.</p>
          </article>
          <article className="feature featureWide paleFeature">
            <div className="featureNumber">04</div>
            <div>
              <h3>Style the entire block.</h3>
              <p>Control TeX size, line spacing, padding, vertical placement, text, background, and border—without routing block styling through Office fills.</p>
            </div>
            <div className="swatches" aria-label="Block styling controls">
              <span style={{ background: "#071738" }} /><span style={{ background: "#08a6aa" }} />
              <span style={{ background: "#ffb11b" }} /><span style={{ background: "#f0dfff" }} />
              <b>Top</b><b>Middle</b><b>Bottom</b>
            </div>
          </article>
        </div>
      </section>

      <section className="hostSection">
        <div className="shell">
          <div className="sectionIntro lightIntro">
            <span className="kicker">Two hosts. One source model.</span>
            <h2>Right at home in Word<br />and PowerPoint.</h2>
          </div>
          <div className="hostCards">
            <article className="hostCard">
              <div className="hostHead"><span className="hostLogo word">W</span><div><small>Microsoft</small><h3>Word</h3></div></div>
              <p>Write in the text stream or build a full-width content block.</p>
              <ul>
                <li><i>✓</i> Inline formulas with exact baselines</li>
                <li><i>✓</i> Display and numbered equations</li>
                <li><i>✓</i> Fixed blocks, inline or floating</li>
                <li><i>✓</i> Copy and paste mixed LaTeX</li>
              </ul>
              <div className="hostSample wordSample"><span>Einstein&apos;s theory</span><b>E = mc²</b><span>changed physics.</span></div>
            </article>
            <article className="hostCard">
              <div className="hostHead"><span className="hostLogo powerpoint">P</span><div><small>Microsoft</small><h3>PowerPoint</h3></div></div>
              <p>Place a free-standing TeX block anywhere on the slide.</p>
              <ul>
                <li><i>✓</i> Exact typesetting width</li>
                <li><i>✓</i> Native position and rotation</li>
                <li><i>✓</i> Reflow after every size change</li>
                <li><i>✓</i> Persistent per-block styling</li>
              </ul>
              <div className="hostSample slideSample"><div /><b>∫<sub>a</sub><sup>b</sup> f(x) dx</b><span>Editable slide object</span></div>
            </article>
          </div>
        </div>
      </section>

      <section className="section shell workflow">
        <div className="sectionIntro centered">
          <span className="kicker">How it works</span>
          <h2>Author once. Keep every layer.</h2>
        </div>
        <div className="steps">
          <article><span>1</span><div className="stepVisual typeVisual">$ E = mc^2 $<i /></div><h3>Write LaTeX</h3><p>Use the source you already know.</p></article>
          <article><span>2</span><div className="stepVisual renderVisual"><i>∫</i><b>StemTeX</b></div><h3>Render precisely</h3><p>A real TeX engine produces portable SVG.</p></article>
          <article><span>3</span><div className="stepVisual placeVisual"><div><i /><b>E = mc²</b></div></div><h3>Work in Office</h3><p>Edit, move, resize, search, and share.</p></article>
        </div>
      </section>

      <section className="downloadSection">
        <div className="shell downloadCard">
          <img src="/brand-icon.png" width="96" height="96" alt="LaTeX Blocks" />
          <div><span className="kicker">Ready to typeset?</span><h2>Make LaTeX part of your Office workflow.</h2><p>One self-contained installer for 64-bit Word and PowerPoint on Windows.</p></div>
          <a className="button downloadButton" href={downloadUrl}>Get LaTeX Blocks <span>↓</span></a>
        </div>
      </section>

      <footer className="footer shell">
        <a className="brand" href="#top"><img src="/brand-icon.png" alt="" width="34" height="34" /><span>LaTeX Blocks</span></a>
        <p>Editable LaTeX for Microsoft Office.</p>
        <div><a href={repositoryUrl}>GitHub</a><a href={`${repositoryUrl}#documentation`}>Documentation</a><a href={`${repositoryUrl}/issues`}>Issues</a></div>
      </footer>
    </main>
  );
}
