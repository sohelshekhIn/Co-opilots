import {AbsoluteFill, Easing, Img, OffthreadVideo, Series, interpolate, staticFile, useCurrentFrame} from 'remotion';

/**
 * Demo 1 — InsideOPS answer.
 * Tone: calm + sharp. Personality: Corporate. One message: it answers from
 * InsideOPS and shows the page.
 * Easing: cubic-bezier(0.22, 1, 0.36, 1). Exits don't travel.
 * Base unit 0.4s. Crossfade only (linear opacity so the overlap sums to 1).
 * Stagger 2 frames (60ms). Travel ≤ 16px. Scale 0.96→1. No overshoot.
 * The mic on the raw recording is left out: this cut would splice it, and
 * the deck is talked over live.
 */

export const FPS = 30;
export const WIDTH = 1920;
export const HEIGHT = 1080;
const FADE = 12;
const UNIT = 12;

const TITLE = 108;
const ASK = 132;
const STATUS = 84;
// Answer appears at source frame 1049. Play it 1x, then hold a still.
// endAt is an absolute source frame: the clip stays mounted for endAt-startFrom
// frames, so that difference must equal the frames we actually play.
const VIDEO_PART = 133;
const HOLD_PART = 72;
const ANSWER = VIDEO_PART + HOLD_PART;

export const DURATION = TITLE + ASK + STATUS + ANSWER - FADE * 3;

const PAPER = '#fafaf8';
const INK = '#0a0a0a';
const MUTED = '#737373';
const LINE = '#d4d4d2';
const ACCENT = '#002FA7';
const FONT = '"Segoe UI", sans-serif';

const easeOut = Easing.bezier(0.22, 1, 0.36, 1);

const ASK_FROM = 66;
const ASK_TO = ASK_FROM + ASK;
const STATUS_FROM = ASK_TO;
const STATUS_TO = STATUS_FROM + STATUS;
const ANSWER_FROM = 1049;
const ANSWER_TO = ANSWER_FROM + VIDEO_PART;

const PANEL_H = 880;
const PANEL_W = Math.round(PANEL_H * (634 / 884));
const PANEL_X = 88;
const PANEL_Y = (HEIGHT - PANEL_H) / 2;

const rise = (frame, delay) => {
  const t = frame - delay;
  const opacity = interpolate(t, [0, UNIT], [0, 1], {
    easing: easeOut,
    extrapolateLeft: 'clamp',
    extrapolateRight: 'clamp',
  });
  const y = interpolate(t, [0, UNIT], [16, 0], {
    easing: easeOut,
    extrapolateLeft: 'clamp',
    extrapolateRight: 'clamp',
  });
  return {opacity, transform: `translateY(${y}px)`};
};

const Shot = ({frames, fadeIn, fadeOut, children}) => {
  const frame = useCurrentFrame();
  const fadeInOpacity = fadeIn ? interpolate(frame, [0, FADE], [0, 1], {extrapolateLeft: 'clamp', extrapolateRight: 'clamp'}) : 1;
  const fadeOutOpacity = fadeOut
    ? interpolate(frame, [frames - FADE, frames], [1, 0], {extrapolateLeft: 'clamp', extrapolateRight: 'clamp'})
    : 1;
  return <AbsoluteFill style={{opacity: fadeInOpacity * fadeOutOpacity}}>{children}</AbsoluteFill>;
};

const Tag = ({n, label}) => (
  <div
    style={{
      display: 'flex',
      alignItems: 'center',
      gap: 12,
      fontFamily: FONT,
      fontSize: 15,
      fontWeight: 600,
      letterSpacing: '0.16em',
      textTransform: 'uppercase',
      color: MUTED,
    }}
  >
    <span
      style={{
        border: `1px solid ${LINE}`,
        color: INK,
        padding: '4px 7px',
        letterSpacing: '0.08em',
        lineHeight: 1,
      }}
    >
      {n}
    </span>
    {label}
  </div>
);

const Title = () => {
  const frame = useCurrentFrame();
  return (
    <AbsoluteFill style={{background: PAPER, padding: '92px 120px', fontFamily: FONT}}>
      <div style={{fontSize: 14, fontWeight: 600, letterSpacing: '0.18em', textTransform: 'uppercase', color: MUTED, ...rise(frame, 0)}}>
        OPS Copilot Assistant
      </div>
      <div
        style={{
          marginTop: 36,
          maxWidth: 760,
          fontSize: 68,
          fontWeight: 300,
          letterSpacing: '-0.03em',
          lineHeight: 1.05,
          color: INK,
          ...rise(frame, 6),
        }}
      >
        First, it answers from InsideOPS and shows the source.
      </div>
      <div style={{display: 'flex', gap: 56, marginTop: 72, alignItems: 'stretch'}}>
        <div style={{flex: 1, ...rise(frame, 20)}}>
          <Tag n="01" label="The situation" />
          <div style={{marginTop: 22, fontSize: 36, fontWeight: 400, letterSpacing: '-0.02em', lineHeight: 1.2, color: INK}}>
            Someone asks what flexible work means.
          </div>
          <div style={{marginTop: 14, fontSize: 20, lineHeight: 1.4, color: MUTED}}>A policy question. Nothing on the PC changes.</div>
        </div>
        <div style={{width: 1, background: LINE}} />
        <div style={{flex: 1, ...rise(frame, 22)}}>
          <Tag n="02" label="Watch for" />
          <div style={{marginTop: 22, fontSize: 36, fontWeight: 400, letterSpacing: '-0.02em', lineHeight: 1.2, color: ACCENT}}>
            The cited page.
          </div>
          <div style={{marginTop: 14, fontSize: 20, lineHeight: 1.4, color: MUTED}}>The answer links the InsideOPS page it used.</div>
        </div>
      </div>
    </AbsoluteFill>
  );
};

const Screen = ({startFrom, endAt, playbackRate, style}) => (
  <OffthreadVideo
    src={staticFile('demo1.mp4')}
    startFrom={startFrom}
    endAt={endAt}
    playbackRate={playbackRate}
    muted
    style={style}
  />
);

const Panel = ({startFrom, endAt, playbackRate = 1, enter, hold, children}) => {
  const frame = useCurrentFrame();
  const scale = enter
    ? interpolate(frame, [0, UNIT], [0.96, 1], {
        easing: easeOut,
        extrapolateLeft: 'clamp',
        extrapolateRight: 'clamp',
      })
    : 1;
  const videoStyle = {position: 'absolute', left: 0, top: 0, width: PANEL_W, height: PANEL_H, objectFit: 'fill'};
  // Subtitle row only (meeting name, sender). Numbers above it stay sharp.
  const bandTop = Math.round(PANEL_H * 0.169);
  const bandH = Math.round(PANEL_H * 0.02);
  const media = (style) =>
    hold ? (
      <Img src={staticFile('answer-hold.png')} style={style} />
    ) : (
      <Screen startFrom={startFrom} endAt={endAt} playbackRate={playbackRate} style={style} />
    );
  return (
    <div
      style={{
        position: 'absolute',
        left: PANEL_X,
        top: PANEL_Y,
        width: PANEL_W,
        height: PANEL_H,
        transform: `scale(${scale})`,
        transformOrigin: '50% 50%',
        borderRadius: 12,
        overflow: 'hidden',
        background: '#fff',
        boxShadow: '0 18px 50px rgba(10,10,10,0.10)',
      }}
    >
      {media(videoStyle)}
      <div style={{position: 'absolute', left: 0, top: bandTop, width: PANEL_W, height: bandH, overflow: 'hidden'}}>
        {media({...videoStyle, top: -bandTop, filter: 'blur(5px)'})}
        <div style={{position: 'absolute', inset: 0, background: 'rgba(244,244,242,0.78)'}} />
      </div>
      {children}
    </div>
  );
};

const Rail = ({n, label, title, body, color}) => (
  <div
    style={{
      position: 'absolute',
      left: PANEL_X + PANEL_W + 64,
      top: PANEL_Y,
      width: WIDTH - (PANEL_X + PANEL_W + 64) - 88,
      height: PANEL_H,
      display: 'flex',
      flexDirection: 'column',
      justifyContent: 'center',
      fontFamily: FONT,
    }}
  >
    <div style={{position: 'absolute', left: -32, top: 80, bottom: 80, width: 1, background: LINE}} />
    <Tag n={n} label={label} />
    <div style={{marginTop: 22, fontSize: 48, fontWeight: 400, letterSpacing: '-0.025em', lineHeight: 1.15, color}}>{title}</div>
    <div style={{marginTop: 16, maxWidth: 460, fontSize: 22, lineHeight: 1.4, color: MUTED}}>{body}</div>
  </div>
);

const Stage = ({children}) => (
  <AbsoluteFill style={{background: PAPER}}>
    <div
      style={{
        position: 'absolute',
        top: 36,
        left: 88,
        fontFamily: FONT,
        fontSize: 13,
        fontWeight: 600,
        letterSpacing: '0.18em',
        textTransform: 'uppercase',
        color: MUTED,
      }}
    >
      OPS Copilot Assistant
    </div>
    <div
      style={{
        position: 'absolute',
        top: 36,
        right: 88,
        fontFamily: FONT,
        fontSize: 13,
        fontWeight: 600,
        letterSpacing: '0.18em',
        textTransform: 'uppercase',
        color: MUTED,
      }}
    >
      Demo 1
    </div>
    {children}
  </AbsoluteFill>
);

const Mark = () => {
  const frame = useCurrentFrame();
  const scaleY = interpolate(frame, [36, 48], [0, 1], {
    easing: easeOut,
    extrapolateLeft: 'clamp',
    extrapolateRight: 'clamp',
  });
  return (
    <div
      style={{
        position: 'absolute',
        left: PANEL_W * 0.145,
        top: PANEL_H * 0.585,
        width: 3,
        height: PANEL_H * 0.19,
        background: ACCENT,
        transform: `scaleY(${scaleY})`,
        transformOrigin: 'top center',
      }}
    />
  );
};

const AnswerBody = () => {
  const frame = useCurrentFrame();
  return (
    <Shot frames={ANSWER} fadeIn fadeOut={false}>
      <Stage>
        <Panel startFrom={ANSWER_FROM} endAt={ANSWER_TO} hold={frame >= VIDEO_PART}>
          <Mark />
        </Panel>
        <Rail
          n="02"
          label="Watch for"
          title="The cited page."
          body="The answer links the InsideOPS page it used."
          color={ACCENT}
        />
      </Stage>
    </Shot>
  );
};

export const Demo1 = () => {
  return (
    <AbsoluteFill style={{background: PAPER}}>
      <Series>
        <Series.Sequence durationInFrames={TITLE}>
          <Shot frames={TITLE} fadeIn={false} fadeOut>
            <Title />
          </Shot>
        </Series.Sequence>
        <Series.Sequence durationInFrames={ASK} offset={-FADE}>
          <Shot frames={ASK} fadeIn fadeOut>
            <Stage>
              <Panel startFrom={ASK_FROM} endAt={ASK_TO} enter>
                {null}
              </Panel>
              <Rail
                n="01"
                label="The situation"
                title="Someone asks what flexible work means."
                body="A policy question. Nothing on the PC changes."
                color={INK}
              />
            </Stage>
          </Shot>
        </Series.Sequence>
        <Series.Sequence durationInFrames={STATUS} offset={-FADE}>
          <Shot frames={STATUS} fadeIn fadeOut>
            <Stage>
              <Panel startFrom={STATUS_FROM} endAt={STATUS_TO} />
              <Rail
                n="02"
                label="Watch for"
                title="Reading InsideOPS."
                body="The status names the source before the answer arrives."
                color={ACCENT}
              />
            </Stage>
          </Shot>
        </Series.Sequence>
        <Series.Sequence durationInFrames={ANSWER} offset={-FADE}>
          <AnswerBody />
        </Series.Sequence>
      </Series>
    </AbsoluteFill>
  );
};
