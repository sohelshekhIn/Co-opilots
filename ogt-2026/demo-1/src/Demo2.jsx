import {AbsoluteFill, Easing, Img, OffthreadVideo, Series, interpolate, staticFile, useCurrentFrame} from 'remotion';

/**
 * Demo 2 — meeting brief.
 * Same language as Demo 1. Tone: calm + sharp. Personality: Corporate.
 * One message: it prepares you for the call and shows each step.
 * Easing: cubic-bezier(0.22, 1, 0.36, 1). Base unit 0.4s. Crossfade only.
 * Stagger 2 frames. Travel ≤ 16px. Scale 0.96→1. No overshoot.
 * The long file search is the cut. Mic left out; the deck is talked over live.
 * endAt is an absolute source frame. The clip stays mounted for endAt-startFrom.
 */

export const FPS = 30;
export const WIDTH = 1920;
export const HEIGHT = 1080;
const FADE = 12;
const UNIT = 12;

const TITLE = 108;
const ASK = 180;
const STEPS = 156;
const VIDEO_PART = 180;
const HOLD_PART = 72;
const BRIEF = VIDEO_PART + HOLD_PART;

export const DURATION = TITLE + ASK + STEPS + BRIEF - FADE * 3;

const PAPER = '#fafaf8';
const INK = '#0a0a0a';
const MUTED = '#737373';
const LINE = '#d4d4d2';
const ACCENT = '#002FA7';
const FONT = '"Segoe UI", sans-serif';

const easeOut = Easing.bezier(0.22, 1, 0.36, 1);

const ASK_FROM = 36 * FPS;
const ASK_TO = ASK_FROM + ASK;
const STEPS_FROM = 90 * FPS;
const STEPS_TO = STEPS_FROM + STEPS;
const BRIEF_FROM = 100 * FPS;
const BRIEF_TO = BRIEF_FROM + VIDEO_PART;

const SRC_W = 1920;
const SRC_H = 1032;
const STAGE_H = 860;
const STAGE_W = Math.round(STAGE_H * (SRC_W / SRC_H));
const STAGE_X = Math.round((WIDTH - STAGE_W) / 2);
const STAGE_Y = 56;

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
    <span style={{border: `1px solid ${LINE}`, color: INK, padding: '4px 7px', letterSpacing: '0.08em', lineHeight: 1}}>{n}</span>
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
          maxWidth: 780,
          fontSize: 64,
          fontWeight: 300,
          letterSpacing: '-0.03em',
          lineHeight: 1.05,
          color: INK,
          ...rise(frame, 6),
        }}
      >
        Next, it prepares you for the call and shows each step.
      </div>
      <div style={{display: 'flex', gap: 56, marginTop: 72, alignItems: 'stretch'}}>
        <div style={{flex: 1, ...rise(frame, 20)}}>
          <Tag n="01" label="The situation" />
          <div style={{marginTop: 22, fontSize: 36, fontWeight: 400, letterSpacing: '-0.02em', lineHeight: 1.2, color: INK}}>
            A call is coming up. What should you know?
          </div>
          <div style={{marginTop: 14, fontSize: 20, lineHeight: 1.4, color: MUTED}}>It looks through the files for that meeting.</div>
        </div>
        <div style={{width: 1, background: LINE}} />
        <div style={{flex: 1, ...rise(frame, 22)}}>
          <Tag n="02" label="Watch for" />
          <div style={{marginTop: 22, fontSize: 36, fontWeight: 400, letterSpacing: '-0.02em', lineHeight: 1.2, color: ACCENT}}>
            Each step, then the brief.
          </div>
          <div style={{marginTop: 14, fontSize: 20, lineHeight: 1.4, color: MUTED}}>The Stop button stays on screen.</div>
        </div>
      </div>
    </AbsoluteFill>
  );
};

const Screen = ({startFrom, endAt, style}) => (
  <OffthreadVideo src={staticFile('demo2.mp4')} startFrom={startFrom} endAt={endAt} muted style={style} />
);

const Stage = ({startFrom, endAt, hold, enter, children}) => {
  const frame = useCurrentFrame();
  const scale = enter
    ? interpolate(frame, [0, UNIT], [0.96, 1], {
        easing: easeOut,
        extrapolateLeft: 'clamp',
        extrapolateRight: 'clamp',
      })
    : 1;
  const mediaStyle = {position: 'absolute', left: 0, top: 0, width: STAGE_W, height: STAGE_H, objectFit: 'fill'};
  const bandTop = Math.round(STAGE_H * 0.088);
  const bandH = Math.round(STAGE_H * 0.028);
  const media = (style) =>
    hold ? (
      <Img src={staticFile('demo2-hold.png')} style={style} />
    ) : (
      <Screen startFrom={startFrom} endAt={endAt} style={style} />
    );
  return (
    <AbsoluteFill style={{background: PAPER}}>
      <div
        style={{
          position: 'absolute',
          top: 22,
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
          top: 22,
          right: 88,
          fontFamily: FONT,
          fontSize: 13,
          fontWeight: 600,
          letterSpacing: '0.18em',
          textTransform: 'uppercase',
          color: MUTED,
        }}
      >
        Demo 2
      </div>
      <div
        style={{
          position: 'absolute',
          left: STAGE_X,
          top: STAGE_Y,
          width: STAGE_W,
          height: STAGE_H,
          transform: `scale(${scale})`,
          transformOrigin: '50% 50%',
          borderRadius: 12,
          overflow: 'hidden',
          background: '#fff',
          boxShadow: '0 18px 50px rgba(10,10,10,0.10)',
        }}
      >
        {media(mediaStyle)}
        <div style={{position: 'absolute', left: 0, top: bandTop, width: STAGE_W, height: bandH, overflow: 'hidden'}}>
          {media({...mediaStyle, top: -bandTop, filter: 'blur(5px)'})}
          <div style={{position: 'absolute', inset: 0, background: 'rgba(244,244,242,0.78)'}} />
        </div>
        {children}
      </div>
    </AbsoluteFill>
  );
};

const Caption = ({n, label, line, color}) => (
  <div
    style={{
      position: 'absolute',
      left: STAGE_X,
      right: STAGE_X,
      bottom: 36,
      display: 'flex',
      alignItems: 'baseline',
      gap: 28,
      fontFamily: FONT,
    }}
  >
    <Tag n={n} label={label} />
    <div style={{fontSize: 28, fontWeight: 400, letterSpacing: '-0.02em', color, lineHeight: 1.2}}>{line}</div>
  </div>
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
        left: STAGE_W * 0.235,
        top: STAGE_H * 0.22,
        width: 3,
        height: STAGE_H * 0.5,
        background: ACCENT,
        transform: `scaleY(${scaleY})`,
        transformOrigin: 'top center',
      }}
    />
  );
};

const BriefBody = () => {
  const frame = useCurrentFrame();
  return (
    <Shot frames={BRIEF} fadeIn fadeOut={false}>
      <Stage startFrom={BRIEF_FROM} endAt={BRIEF_TO} hold={frame >= VIDEO_PART}>
        <Mark />
      </Stage>
      <Caption n="02" label="Watch for" line="The brief for the call." color={ACCENT} />
    </Shot>
  );
};

export const Demo2 = () => {
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
            <Stage startFrom={ASK_FROM} endAt={ASK_TO} enter />
            <Caption n="01" label="The situation" line="A call is coming up. What should you know?" color={INK} />
          </Shot>
        </Series.Sequence>
        <Series.Sequence durationInFrames={STEPS} offset={-FADE}>
          <Shot frames={STEPS} fadeIn fadeOut>
            <Stage startFrom={STEPS_FROM} endAt={STEPS_TO} />
            <Caption n="02" label="Watch for" line="Each step stays visible. So does Stop." color={ACCENT} />
          </Shot>
        </Series.Sequence>
        <Series.Sequence durationInFrames={BRIEF} offset={-FADE}>
          <BriefBody />
        </Series.Sequence>
      </Series>
    </AbsoluteFill>
  );
};
